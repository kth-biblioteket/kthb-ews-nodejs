require('dotenv').config()
require("datejs")

const EWS = require('node-ews');

const ewsConfig = {
  host: process.env.EWS_HOST,
  username: process.env.EWS_USER,
  password: process.env.EWS_PASS
};

const ews = new EWS(ewsConfig);

const ewsFunction = 'FindItem';

let StartDate = new Date('2020-01-01 00:00');
let EndDate = new Date('2020-12-31 23:59');
StartDate = StartDate.toISOString();
EndDate = EndDate.toISOString();

const EmailAddress = 'ece-biblioteket@ug.kth.se';

const ewsArgs = {
    'attributes': {
      'Traversal': 'Shallow'
    },
    'ItemShape': {
      'BaseShape': 'AllProperties'
    },
    'CalendarView ': {
      'attributes': {
        'StartDate': StartDate,
        'EndDate': EndDate
      }
    },
    'ParentFolderIds' : {
      'DistinguishedFolderId': {
        'attributes': {
          'Id': 'calendar'
        },
        'Mailbox': {
          'EmailAddress': EmailAddress
        }
      }
    }
};

function getcalendarevents(ewsFunction, ewsArgs) {
  ews.run(ewsFunction, ewsArgs)
    .then(result => {
      console.log(JSON.stringify(result));
    })
    .catch(err => {
      console.log(err.stack);
    });
}

getcalendarevents(ewsFunction, ewsArgs);