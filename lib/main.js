"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const core = __importStar(require("@actions/core"));
const graph_1 = require("@pnp/graph");
const nodejs_1 = require("@pnp/nodejs");
const date_fns_1 = require("date-fns");
function run() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const tenantName = process.env.TENANT_NAME;
            const appId = process.env.APP_ID;
            const appSecret = process.env.APP_SECRET;
            const subject = core.getInput('subject');
            const body = core.getInput('body');
            const emailAddress = core.getInput('emailaddress');
            let startTime = core.getInput('start');
            let endTime = core.getInput('end');
            if (startTime == null) {
                let nextDay = date_fns_1.format(date_fns_1.addBusinessDays(new Date(), 1), 'yyyy-MM-dd');
                core.debug(nextDay);
                startTime = `${nextDay}T12:00:00`;
                endTime = `${nextDay}T13:00:00`;
            }
            console.log("Setting up graph...");
            graph_1.graph.setup({
                graph: {
                    fetchClientFactory: () => {
                        return new nodejs_1.AdalFetchClient(tenantName, appId, appSecret);
                    },
                },
            });
            console.log("Creating event...");
            let e = {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": body
                },
                "start": {
                    "dateTime": startTime,
                    "timeZone": "GMT Standard Time"
                },
                "end": {
                    "dateTime": endTime,
                    "timeZone": "GMT Standard Time"
                },
                "location": {
                    "displayName": "Office"
                }
            };
            yield graph_1.graph.users.getById(emailAddress).calendar.events.add(e);
            console.log("Event created.");
        }
        catch (error) {
            core.setFailed(error.message);
        }
    });
}
run();
