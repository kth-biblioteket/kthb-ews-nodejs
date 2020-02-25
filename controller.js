const EWS = require('node-ews');
const ewsConfig = {
    host: process.env.EWS_HOST,
    username: process.env.EWS_USER,
    password: process.env.EWS_PASS
};

let ews = new EWS(ewsConfig);

exports.getCalendarEvents = async function (req, res) {
    
  const ewsFunction = 'FindItem';
  
  const EmailAddress = req.params.emailaddress;

  let StartDate
  let EndDate
  try {
    StartDate = new Date(req.query.startdate);
    EndDate = new Date(req.query.enddate);
    
    StartDate = StartDate.toISOString();
    EndDate = EndDate.toISOString();
  }
  catch(err) {
    return res.status(500).json({'error: ' : err.toString()});
  }
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
  

    ews.run(ewsFunction, ewsArgs)
    .then(result => {
        res.json(result);
    })
    .catch(err => {
      console.log(err.stack);
    });

    ews = null;
  
};