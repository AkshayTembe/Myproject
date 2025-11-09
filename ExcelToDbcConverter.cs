// ExcelToDbcConverter.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using DbcParserLib.Excel.Models;
using DbcParserLib.Excel.Observers;
using DbcParserLib.Excel.SheetParsers;
using DbcParserLib;
using DbcParserLib.Model;

namespace DbcParserLib.Excel
{
    public static class ExcelToDbcConverter
    {
        /// <summary>
        /// Parse the excel file at excelPath and write .dbc to outputDbcPath.
        /// Returns an ExcelParsingResult with errors/warnings and the Dbc (if success).
        /// </summary>
        public static ExcelParsingResult ParseAndWrite(string excelPath, string outputDbcPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath)) throw new ArgumentNullException(nameof(excelPath));
            if (string.IsNullOrWhiteSpace(outputDbcPath)) throw new ArgumentNullException(nameof(outputDbcPath));
            if (!File.Exists(excelPath)) return ExcelParsingResult.CreateFailure(new List<string> { $"Excel file '{excelPath}' not found" });

            var observer = new SimpleExcelObserver();
            var data = new ExcelDbcData();

            // Ensure EPPlus license context set in consumer app if required:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                // instantiate sheet parsers in a sensible order
                var parsers = new ISheetParser[]
                {
                    new NodesSheetParser(),
                    new MessagesSheetParser(),
                    new ValueTablesSheetParser(),
                    new SignalsSheetParser(),
                    new EnvironmentVariablesSheetParser(),
                    new BaDefSheetParser(),
                    new BaSheetParser(),
                    new ExtraTransmittersSheetParser()
                };

                // run each parser (if sheet exists)
                foreach (var parser in parsers)
                {
                    observer.CurrentSheet = parser.SheetName;

                    var ws = ExcelHelpers.GetWorksheet(package, parser.SheetName);
                    if (ws == null)
                    {
                        if (parser.IsRequired)
                        {
                            observer.SheetNotFound(parser.SheetName);
                        }
                        // optional sheet missing: continue
                        continue;
                    }

                    // Let parser validate required columns and parse
                    var success = parser.Parse(package, data, observer);
                    // Continue parsing other sheets even on failure to accumulate errors
                }
            }

            // Merge parsed signals into messages
            foreach (var kv in data.Signals)
            {
                var msgId = kv.Key;
                if (!data.Messages.ContainsKey(msgId))
                {
                    // signal references missing message — add warning and skip
                    observer.Warning($"Signals present for message 0x{msgId:X} but message not found");
                    continue;
                }

                var message = data.Messages[msgId];
                foreach (var s in kv.Value)
                {
                    s.Parent = message;
                    message.Signals.Add(s);
                }
            }

            // Apply ExtraTransmitters
            foreach (var kv in data.ExtraTransmitters)
            {
                var msgId = kv.Key;
                if (data.Messages.TryGetValue(msgId, out var msg))
                {
                    msg.AdditionalTransmitters = kv.Value;
                }
                else
                {
                    observer.Warning($"ExtraTransmitters present for unknown message 0x{msgId:X}");
                }
            }

            // Apply property assignments (BA_ entries)
            foreach (var pa in data.PropertyAssignments)
            {
                try
                {
                    // find definition for scope+property
                    if (!data.CustomPropertyDefinitions.TryGetValue(pa.Scope, out var defsForScope) ||
                        !defsForScope.TryGetValue(pa.PropertyName, out var def))
                    {
                        observer.PropertyNotFound(pa.PropertyName);
                        continue;
                    }

                    var customProp = new CustomProperty(def);
                    if (!customProp.SetCustomPropertyValue(pa.Value, pa.IsNumeric))
                    {
                        observer.PropertyValueInvalid(pa.PropertyName, pa.Value);
                        continue;
                    }

                    switch (pa.Scope)
                    {
                        case CustomPropertyObjectType.Global:
                            data.GlobalProperties[def.Name] = customProp;
                            break;

                        case CustomPropertyObjectType.Node:
                            var node = data.Nodes.FirstOrDefault(n => n.Name == pa.ScopeIdentifier);
                            if (node == null)
                            {
                                observer.NodeReferenceNotFound(pa.ScopeIdentifier, $"BA assignment for '{pa.PropertyName}'");
                                continue;
                            }
                            node.CustomProperties[def.Name] = customProp;
                            break;

                        case CustomPropertyObjectType.Message:
                            if (!ExcelHelpers.TryParseHexId(pa.ScopeIdentifier, out var mid) || !data.Messages.TryGetValue(mid, out var message))
                            {
                                observer.MessageReferenceNotFound(pa.ScopeIdentifier, $"BA assignment for '{pa.PropertyName}'");
                                continue;
                            }
                            message.CustomProperties[def.Name] = customProp;
                            break;

                        case CustomPropertyObjectType.Signal:
                            // scope identifier expected "msgId:signalName"
                            var parts = pa.ScopeIdentifier?.Split(':') ?? new string[0];
                            if (parts.Length != 2)
                            {
                                observer.Warning($"Invalid signal scope identifier format: '{pa.ScopeIdentifier}'. Expected 'messageId:signalName'");
                                continue;
                            }

                            if (!ExcelHelpers.TryParseHexId(parts[0], out var smid) || !data.Messages.TryGetValue(smid, out var smsg))
                            {
                                observer.MessageReferenceNotFound(parts[0], $"BA assignment for '{pa.PropertyName}'");
                                continue;
                            }

                            var sig = smsg.Signals.FirstOrDefault(s => s.Name == parts[1]);
                            if (sig == null)
                            {
                                observer.SignalReferenceNotFound(parts[0], parts[1], $"BA assignment for '{pa.PropertyName}'");
                                continue;
                            }

                            sig.CustomProperties[def.Name] = customProp;
                            break;

                        case CustomPropertyObjectType.Environment:
                            if (!data.EnvironmentVariables.TryGetValue(pa.ScopeIdentifier, out var ev))
                            {
                                observer.Warning($"Environment variable '{pa.ScopeIdentifier}' not found for BA assignment");
                                continue;
                            }
                            ev.CustomProperties[def.Name] = customProp;
                            break;
                    }
                }
                catch (Exception ex)
                {
                    observer.UnexpectedError($"Applying BA assignment '{pa.PropertyName}' failed: {ex.Message}");
                }
            }

            // Apply comments stored in ExcelDbcData.Comments (if any)
            foreach (var c in data.Comments)
            {
                var typeNorm = ValidationHelpers.NormalizeCommentType(c.Type);
                if (typeNorm == "BO")
                {
                    if (ExcelHelpers.TryParseHexId(c.Scope, out var mid) && data.Messages.TryGetValue(mid, out var mm))
                    {
                        mm.Comment = c.Comment;
                    }
                    else
                    {
                        observer.CommentScopeInvalid(c.Scope);
                    }
                }
                else if (typeNorm == "SG")
                {
                    // scope expected "msgId:signalName"
                    var parts = c.Scope?.Split(':') ?? new string[0];
                    if (parts.Length != 2)
                    {
                        observer.CommentScopeInvalid(c.Scope);
                        continue;
                    }
                    if (ExcelHelpers.TryParseHexId(parts[0], out var mid2) && data.Messages.TryGetValue(mid2, out var mm2))
                    {
                        var sig = mm2.Signals.FirstOrDefault(s => s.Name == parts[1]);
                        if (sig != null) sig.Comment = c.Comment;
                        else observer.CommentScopeInvalid(c.Scope);
                    }
                    else observer.CommentScopeInvalid(c.Scope);
                }
                else if (typeNorm == "BU")
                {
                    var node = data.Nodes.FirstOrDefault(n => n.Name == c.Scope);
                    if (node != null) node.Comment = c.Comment;
                    else observer.CommentScopeInvalid(c.Scope);
                }
                else if (typeNorm == "EV")
                {
                    if (data.EnvironmentVariables.TryGetValue(c.Scope, out var ev))
                        ev.Comment = c.Comment;
                    else observer.CommentScopeInvalid(c.Scope);
                }
            }

            // Convert dictionaries / lists into the final Dbc model expected by DbcWriter
            var nodesForDbc = data.Nodes;
            var messagesForDbc = data.Messages.Values;
            var envVarsForDbc = data.EnvironmentVariables.Values;
            var globalPropsForDbc = data.GlobalProperties.Values;
            var namedValueTables = data.NamedValueTables;
            var customPropDefs = data.CustomPropertyDefinitions;

            // Build Dbc
            var dbc = new Dbc(
                nodesForDbc,
                messagesForDbc,
                envVarsForDbc,
                globalPropsForDbc,
                namedValueTables,
                customPropDefs);

            try
            {
                DbcWriter.WriteToPath(outputDbcPath, dbc);
            }
            catch (Exception ex)
            {
                observer.UnexpectedError($"Writing DBC failed: {ex.Message}");
                return ExcelParsingResult.CreateFailure(observer.GetErrorList());
            }

            // Prepare parsing result
            if (observer.HasErrors())
            {
                return ExcelParsingResult.CreateFailure(observer.GetErrorList(), observer.GetWarningList());
            }
            else if (observer.HasWarnings())
            {
                return ExcelParsingResult.CreateSuccessWithWarnings(dbc, observer.GetWarningList());
            }
            else
            {
                return ExcelParsingResult.CreateSuccess(dbc);
            }
        }
    }
}
