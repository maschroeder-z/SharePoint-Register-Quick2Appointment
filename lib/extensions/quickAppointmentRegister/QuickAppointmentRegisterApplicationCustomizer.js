var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
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
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
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
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'QuickAppointmentRegisterApplicationCustomizerStrings';
import { SPHttpClient } from '@microsoft/sp-http';
var LOG_SOURCE = 'dev-sky-QuickAppointmentRegister';
var QuickAppointmentRegisterApplicationCustomizer = /** @class */ (function (_super) {
    __extends(QuickAppointmentRegisterApplicationCustomizer, _super);
    function QuickAppointmentRegisterApplicationCustomizer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    QuickAppointmentRegisterApplicationCustomizer.prototype.onInit = function () {
        Log.info(LOG_SOURCE, "Fast appointment register extension loaded.");
        this.specialClientSideExtensions();
        return Promise.resolve();
    };
    QuickAppointmentRegisterApplicationCustomizer.prototype.specialClientSideExtensions = function () {
        var _this = this;
        this.context.application.navigatedEvent.add(this, function () {
            setTimeout(function () {
                _this.extendEventPage();
            }, 800);
        });
    };
    QuickAppointmentRegisterApplicationCustomizer.prototype.extendEventPage = function () {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var container, allEvents, params, listGuid_1, itemID_1, currentUser_1, btnRegister, currentAppointmentEntry_1, propertyContent, head3;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!(location.href.toLowerCase().indexOf("/_layouts/15/event.aspx") !== -1)) return [3 /*break*/, 5];
                        Log.info(LOG_SOURCE, "We are on an event page!");
                        container = document.querySelector("section[data-automation-id='seeAllEvents']");
                        if (!(container !== undefined && container !== null)) return [3 /*break*/, 4];
                        allEvents = container.querySelector("a");
                        Log.info(LOG_SOURCE, "Found container to add our new function!");
                        params = new URLSearchParams(location.search);
                        listGuid_1 = params.get('ListGuid');
                        itemID_1 = params.get('ItemId') !== null ? parseInt(params.get('ItemId'), 10) : 0;
                        currentUser_1 = this.context.pageContext.user;
                        if (!(listGuid_1 !== null && itemID_1 > 0)) return [3 /*break*/, 2];
                        btnRegister = document.createElement("a");
                        if (allEvents !== undefined && allEvents !== null)
                            btnRegister.className = allEvents.className;
                        btnRegister.style.marginLeft = "10px";
                        btnRegister.style.paddingTop = "5px";
                        return [4 /*yield*/, this.loadAppointment(listGuid_1, itemID_1)];
                    case 1:
                        currentAppointmentEntry_1 = _b.sent();
                        if (currentAppointmentEntry_1.ParticipantsPickerId.filter(function (x) { return x.Title === currentUser_1.displayName; }).length > 0) {
                            propertyContent = (_a = container.parentNode) === null || _a === void 0 ? void 0 : _a.firstChild;
                            if (propertyContent !== undefined) {
                                head3 = document.createElement("h3");
                                head3.innerText = strings.HeadRegistered;
                                propertyContent.before(head3);
                            }
                            btnRegister.innerText = strings.BTNUnregister;
                            btnRegister.onclick = function (source) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this.manageUserToAppointment(listGuid_1, itemID_1, currentAppointmentEntry_1, currentUser_1, false)];
                                        case 1:
                                            _a.sent();
                                            Dialog.alert(strings.MSGUnregistered).then(function () {
                                                location.reload();
                                            });
                                            return [2 /*return*/];
                                    }
                                });
                            }); };
                        }
                        else {
                            btnRegister.innerText = strings.BTNRegister;
                            btnRegister.onclick = function (source) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this.manageUserToAppointment(listGuid_1, itemID_1, currentAppointmentEntry_1, currentUser_1, true)];
                                        case 1:
                                            _a.sent();
                                            Dialog.alert(strings.MSGRegistered).then(function () {
                                                location.reload();
                                            });
                                            return [2 /*return*/];
                                    }
                                });
                            }); };
                        }
                        container.appendChild(btnRegister);
                        return [3 /*break*/, 3];
                    case 2:
                        Log.warn(LOG_SOURCE, "Cannot apply register function due to missing data: ListID: ".concat(listGuid_1, ", ItemID: ").concat(itemID_1, "."));
                        _b.label = 3;
                    case 3: return [3 /*break*/, 5];
                    case 4:
                        Log.warn(LOG_SOURCE, "Cannot apply register function due to missing CONTAINER element!");
                        _b.label = 5;
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    QuickAppointmentRegisterApplicationCustomizer.prototype.loadAppointment = function (listGuid, id) {
        var endpoint = "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists/getById('").concat(listGuid, "')/items(").concat(id, ")?$expand=ParticipantsPicker&$select=Id,Title,ParticipantsPicker/ID,ParticipantsPicker/Title");
        return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        })
            .then(function (jsonResponse) {
            return {
                Id: jsonResponse.Id,
                ParticipantsPickerId: jsonResponse.ParticipantsPicker ? jsonResponse.ParticipantsPicker : [],
                Title: jsonResponse.Title
            };
        });
    };
    QuickAppointmentRegisterApplicationCustomizer.prototype.manageUserToAppointment = function (listGuid, id, userList, userToAdd, addNew) {
        return __awaiter(this, void 0, void 0, function () {
            var clientconfig, options, reqUser, userData, body, updateoptions;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        clientconfig = SPHttpClient.configurations.v1;
                        options = {
                            body: JSON.stringify({ 'logonName': userToAdd.loginName })
                        };
                        return [4 /*yield*/, this.context.spHttpClient.post("".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/ensureuser"), clientconfig, options)];
                    case 1:
                        reqUser = _a.sent();
                        return [4 /*yield*/, reqUser.json()];
                    case 2:
                        userData = _a.sent();
                        if (addNew) {
                            userList.ParticipantsPickerId = userList.ParticipantsPickerId.map(function (x) { return x.ID; });
                            userList.ParticipantsPickerId.push(userData.Id);
                        }
                        else {
                            userList.ParticipantsPickerId = userList.ParticipantsPickerId.filter(function (x) { return x.ID !== userData.Id; }).map(function (x) { return x.ID; });
                        }
                        body = JSON.stringify(userList);
                        updateoptions = {
                            headers: {
                                'Accept': 'application/json;odata=nometadata',
                                'Content-type': 'application/json;odata=nometadata',
                                'odata-version': '',
                                'IF-MATCH': '*',
                                'X-HTTP-Method': 'MERGE'
                            },
                            body: body
                        };
                        return [2 /*return*/, this.context.spHttpClient.post("".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists/getById('").concat(listGuid, "')/items(").concat(id, ")"), clientconfig, updateoptions)];
                }
            });
        });
    };
    return QuickAppointmentRegisterApplicationCustomizer;
}(BaseApplicationCustomizer));
export default QuickAppointmentRegisterApplicationCustomizer;
//# sourceMappingURL=QuickAppointmentRegisterApplicationCustomizer.js.map