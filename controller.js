const EWS = require('node-ews');
const ewsConfig = {
    host: process.env.EWS_HOST,
    username: process.env.EWS_USER,
    password: process.env.EWS_PASS
};

exports.getCalendarEvents = async function (req, res) {
    
    const ews = new EWS(ewsConfig);

    const ewsFunction = 'FindItem';
    
    const EmailAddress = req.params.emailaddress;
    
    let StartDate = new Date(req.query.startdate);
    let EndDate = new Date(req.query.enddate);
    
    StartDate = StartDate.toISOString();
    EndDate = EndDate.toISOString();

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
        //res.json(result);
        JSON.stringify(result)
    })
    .catch(err => {
      console.log(err.stack);
    });

    
};