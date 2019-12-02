import * as core from '@actions/core';
import { graph } from "@pnp/graph";
import { AdalFetchClient } from "@pnp/nodejs";
import { addBusinessDays, format } from 'date-fns'

async function run() {
    try {

        const tenantName = process.env.TENANT_NAME;
        const appId = process.env.APP_ID;
        const appSecret = process.env.APP_SECRET;

        const subject = core.getInput('subject');
        const body = core.getInput('body');
        const emailAddress = core.getInput('emailaddress');

        let startTime = core.getInput('start');
        let endTime = core.getInput('end');

        if(startTime.length == 0){
            let nextDay: any = format(addBusinessDays(new Date(), 1), 'yyyy-MM-dd');
            core.debug(nextDay);

            startTime = `${nextDay}T12:00:00`;
            endTime = `${nextDay}T13:00:00`
        }

        console.log("Setting up graph...");

        graph.setup({
            graph: {
                fetchClientFactory: () => {
                    return new AdalFetchClient(tenantName, appId, appSecret);
                },
            },
        });

        console.log("Creating event...");

        let e: any = {
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

        await graph.users.getById(emailAddress).calendar.events.add(e);

        console.log("Event created.");

    } catch (error) {
        core.setFailed(error.message);
    }
}

run();