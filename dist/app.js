/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
(function () {
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#assignRange').click(assignRange);
            $('#copyRange').click(copyRange);
            $('#loadProperty').click(loadProperty);
            $('#singleInputCopy').click(singleInputCopy);
            $('#bindingFromA1Range').click(bindingFromA1Range);
            $('#getAllbdings').click(getAllbdings);
            $('#getBindingData').click(getBindingData);
            $('#getBindingWithOfficeSelect').click(getBindingWithOfficeSelect);
            $('#addHandler').click(addHandler);
        });
    };
    function assignRange() {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                            var sheet, values, range;
                            return __generator(this, function (_a) {
                                /**
                                 * Insert your Excel code here
                                 */
                                console.log('assignRange is clicked');
                                sheet = context.workbook.worksheets.getActiveWorksheet();
                                values = [
                                    ["Type", "Estimate"],
                                    ["Transportation", 1670]
                                ];
                                range = sheet.getRange("A1:B2");
                                // Assign array value to the proxy object's values property.
                                range.values = values;
                                // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
                                return [2 /*return*/, context.sync().then(function () {
                                        console.log("Done");
                                    })];
                            });
                        }); }).catch(function (error) {
                            console.log("Error: " + error);
                            if (error instanceof OfficeExtension.Error) {
                                console.log("Debug info: " + JSON.stringify(error.debugInfo));
                            }
                            // await context.sync();
                        })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    }
    function copyRange() {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log('copyRange is clicked');
                        // Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
                        return [4 /*yield*/, Excel.run(function (ctx) {
                                // Create a proxy object for the range and load the values property
                                var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2").load("values");
                                // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
                                return ctx.sync().then(function () {
                                    // Assign the previously loaded values to the new range proxy object. The values will be updated once the following .then() function is invoked.
                                    ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = range.values;
                                });
                            }).then(function () {
                                console.log("done");
                            }).catch(function (error) {
                                console.log("Error: " + error);
                                if (error instanceof OfficeExtension.Error) {
                                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                                }
                            })];
                    case 1:
                        // Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    }
    function loadProperty() {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log('loadProperty is loaded');
                        return [4 /*yield*/, Excel.run(function (ctx) {
                                var sheetName = "Sheet1";
                                var rangeAddress = "A1:B2";
                                var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
                                myRange.load(["values", "address", "format/*", "format/fill", "entireRow"]);
                                return ctx.sync().then(function () {
                                    console.log(myRange.values);
                                    console.log(myRange.address); //ok
                                    console.log(myRange.format.wrapText); //ok
                                    console.log(myRange.format.fill.color); //ok
                                    console.log(myRange.format.font.color); //not ok as it was not loaded
                                });
                            }).then(function () {
                                console.log("done");
                            }).catch(function (error) {
                                console.log("Error: " + error);
                                if (error instanceof OfficeExtension.Error) {
                                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                                }
                            })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    }
    function singleInputCopy() {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log('singleInputCopy is called');
                        return [4 /*yield*/, Excel.run(function (ctx) {
                                var sheetName = 'Sheet1';
                                var rangeAddress = 'A1:A20';
                                var worksheet = ctx.workbook.worksheets.getItem(sheetName);
                                var range = worksheet.getRange(rangeAddress);
                                // range.values = 'Due Date';
                                range.load('text');
                                return ctx.sync().then(function () {
                                    console.log(range.text);
                                });
                            }).catch(function (error) {
                                console.log("Error: " + error);
                                if (error instanceof OfficeExtension.Error) {
                                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                                }
                            })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    }
    function bindingFromA1Range() {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                Office.context.document.bindings.addFromNamedItemAsync("A1:A3", Office.BindingType.Matrix, { id: "MyCities" }, function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        write('Error: ' + asyncResult.error.message);
                    }
                    else {
                        // Write data to the new binding.
                        Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" }, function (asyncResult) {
                            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                                console.log(asyncResult);
                                write('Control bound. Binding.id: '
                                    + asyncResult.value.id + ' Binding.type: ' + asyncResult.value.type);
                            }
                            else if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                write('Error: ' + asyncResult.error.message);
                            }
                        });
                    }
                });
                return [2 /*return*/];
            });
        });
    }
    function getAllbdings() {
        Office.context.document.bindings.getAllAsync(function (asyncResult) {
            var bindingString = '';
            for (var i in asyncResult.value) {
                bindingString += asyncResult.value[i].id + '\n';
            }
            write('Existing bindings: ' + bindingString);
        });
    }
    function getBindingData() {
        Office.context.document.bindings.getByIdAsync('MyCities', function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Action failed. Error: ' + asyncResult.error.message);
            }
            else {
                write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
                var myBinding = asyncResult.value;
                myBinding.getDataAsync(function (asyncResult2) {
                    if (asyncResult2.status == Office.AsyncResultStatus.Failed) {
                        write('Action failed. Error: ' + asyncResult2.error.message);
                    }
                    else {
                        write(asyncResult2.value);
                    }
                });
            }
        });
    }
    function getBindingWithOfficeSelect() {
        Office.select("bindings#MyCities", function onError() { }).getDataAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Action failed. Error: ' + asyncResult.error.message);
            }
            else {
                write(asyncResult.value);
            }
        });
    }
    function addHandler() {
        Office.select("bindings#MyCities").addHandlerAsync(Office.EventType.BindingDataChanged, dataChanged);
    }
    function dataChanged(eventArgs) {
        write('Bound data changed in binding: ' + eventArgs.binding.id);
        getBindingWithOfficeSelect();
    }
    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message + '\n';
    }
})();
//# sourceMappingURL=app.js.map