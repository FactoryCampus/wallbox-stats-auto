'use strict';

const Mailgun = require("mailgun-js");
const request = require("request-promise-native");
const parseCSV = require("csv-parse/lib/sync");
const { DateTime } = require('luxon');
const XLSX = require('xlsx');

const parseDate = val => {
    const parsedDate = DateTime.fromString(val, "dd/MM/yyyy");
    if(parsedDate.isValid) {
        return parsedDate.toJSDate();
    }
    return val;
};

const currentDate = DateTime.utc();
const lastMonthStart = DateTime.utc(currentDate.year, currentDate.month-1);
const lastMonthEnd = DateTime.utc(currentDate.year, currentDate.month-1).endOf('month');
const isLastMonth = date => date.getTime() >= lastMonthStart.valueOf() && date.getTime() <= lastMonthEnd.valueOf();

module.exports = (context, req, res) => {
    const { mgApiKey, mgDomain, targetUrl, targetPassword, reportReceipient, errorReceipient } = context.secrets;
    const mg = Mailgun({ apiKey: mgApiKey, domain: mgDomain });

    // fetch csv
    request({
        url: targetUrl,
        auth: {
            user: "admin",
            pass: targetPassword
        }
    })
        .then(body => parseCSV(body, { columns: true, delimiter: ';', auto_parse: true, auto_parse_date: parseDate }))
        .then(table => table.filter(line => isLastMonth(line.Startdatum)))
        .then(table => XLSX.utils.json_to_sheet(table, { cellDates: true, dateNF: "dd.mm.yyyy" }))
        .then(ws => ({ SheetNames: ['Wallbox'], Sheets: { Wallbox: ws } }))
        .then(wb => XLSX.write(wb, { type: 'buffer' }))
        .then(buf => new mg.Attachment({ data: buf, filename: `wallbox-usage_${lastMonthStart.year}-${lastMonthStart.month}.xlsx` }))
        .then(attach => ({
            from: 'Wallbox Usage <it@factorycampus.de>',
            to: reportReceipient,
            subject: `Wallbox Nutzung ${lastMonthStart.setLocale('de-DE').month} ${lastMonthStart.year}`,
            text: 'Anbei finden Sie den Report über die Aufladungen an der Wallbox im vergangenen Monat.',
            attachment: attach
        }))
        .then(emailData => mg.messages().send(emailData))
        .catch(err => {
            console.error(err);
            mg.messages().send({
                from: 'Wallbox Usage <it@factorycampus.de>',
                to: errorReceipient,
                subject: "💥 Wallbox Cron failed!",
                text: "The following error occurred when trying to send a usage report for the wallbox:\n" + err.toString()
            });
        })
        .then(() => res.end());
};
