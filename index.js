import nodemailer from "nodemailer";
import { writeFileSync } from "fs";
import ics from "ics";

// for the ics file
const eventtitle = 'Vacation';
const eventdescription = '';
const eventstatus = 'BUSY';
const eventlocation = 'Germany';
const eventmethod = 'request';
const eventattendees = { name: 'Max Mustermann', email: 'max@germany.de', rsvp: true, partstat: 'ACCEPTED', role: 'REQ-PARTICIPANT' };
const eventorganizer = { name: 'Max Mustermann', email: 'max@germany.de' };
// year, month, day
const starttime = [2023, 3, 15];
const durationtime = { days: 1 };

// for sending mail
const frommail = 'germany@germany.de';
const tomail = 'max@germany.de; max2@germany.de';
const mailsubject = 'Max Mustermann ICS File';
const hostserver = 'webmail.germanyde';
const hostport = '587';

// authentication
const userdomain = 'd';
const loginname = 'maxm';
const loginpasswort = 'xxxyyyzzz';

// transport of the mail (OWA [Office 365])
var o365Transport = nodemailer.createTransport({
    host: hostserver,
    secure: false,
    port: hostport,
    tls: {
        ciphers: "SSLv3",
        rejectUnauthorized: false,
    },
    auth: {
        domain: userdomain,
        user: loginname,
        pass: loginpasswort,
    },
    debug: true,
    logger:true,
});

// generate ics file
ics.createEvent({
  title: eventtitle,
  description: eventdescription,
  busyStatus: eventstatus,
  location: eventlocation,
  method: eventmethod,
  start: starttime,
  duration: durationtime,
  organizer: eventorganizer,
  attendees: [eventattendees]
}, (error, value) => {
  if (error) {
    console.log(error)
  }
  writeFileSync(`event.ics`, value)
})

// build the mail and send it
async function sendemail() {
    let message = {
        from: frommail,
        to: tomail,
        subject: mailsubject,
        text: '',
        icalEvent: {
            filename: 'event.ics',
            method: eventmethod,
            path: 'event.ics',
            encoding: 'base64'
        }
    }

    o365Transport.sendMail(message, function (error, response) {
        if (error) {
            console.log(error);
        } else {
            console.log("Message sent: " , response);
        }
    })
}

sendemail();