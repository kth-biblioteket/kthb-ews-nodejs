const EWS = require('node-ews');
const ewsConfig = {
    host: process.env.EWS_HOST,
    username: process.env.EWS_USER,
    password: process.env.EWS_PASS
};

let ews = new EWS(ewsConfig);

exports.getItems = async function (req, res) {
    const ewsArgs = {
        'ItemShape': {
            'BaseShape': 'IdOnly'
        },
        'ItemIds': {
            'ItemId': {
                'attributes': {
                    'Id': req.params.itemid
                }
            }
        }
    };

    await ews.run('GetItem', ewsArgs)
    .then(result => {
        res.json(result);
    })
    .catch(err => {
        console.log(err.stack);
    });

}

exports.getCalendarItems = async function (req, res) {
    
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
  
};

exports.createCalendarItems = async function (req, res) {
    
    const event = req.body;

    const ewsFunction = 'CreateItem';
    
    const EmailAddress = req.body.emailaddress;
  
    let location
    let start
    let end
    let subject
    
    subject = req.body.subject

    try {
        start = new Date(req.body.start);
        end = new Date(req.body.end);
      
        start = start.toISOString();
        end = end.toISOString();
    }
    catch(err) {
      return res.status(500).json({'error: ' : err.toString()});
    }
    
    location = req.body.location

    const ewsArgs = {
        'attributes': {
            'SendMeetingInvitations': 'SendToNone'
        },
        'SavedItemFolderId' : {
            'DistinguishedFolderId': {
              'attributes': {
                'Id': 'calendar'
              },
              'Mailbox': {
                'EmailAddress': EmailAddress
              }
            }
        },
        'Items': {
            'CalendarItem': {
                'Subject': subject,
                'Start': start,
                'End': end,
                'Location': location,
            }
        }
    };
    //res.json(ewsArgs);
    ews.run(ewsFunction, ewsArgs)
    .then(result => {
        res.json(result);
    })
    .catch(err => {
    console.log(err.stack);
    });
  };

exports.cancelCalendarItems = async function (req, res) {

    const ewsFunction = 'CreateItem';
    
    const emailaddress = req.body.emailaddress;
    const itemid = req.body.itemid;
  
    let changekey = ''

    let ewsArgs = {
        'ItemShape': {
            'BaseShape': 'IdOnly'
        },
        'ItemIds': {
            'ItemId': {
                'attributes': {
                    'Id': itemid
                }
            }
        }
    };

    let itemresult
    await ews.run('GetItem', ewsArgs)
    .then(result => {
        itemresult = result;
        if(result.ResponseMessages.GetItemResponseMessage.attributes.ResponseClass != 'Error') {
            changekey = result.ResponseMessages.GetItemResponseMessage.Items.CalendarItem.ItemId.attributes.ChangeKey
        }
    })
    .catch(err => {
        console.log(err.stack);
    });

    if(itemresult.ResponseMessages.GetItemResponseMessage.attributes.ResponseClass == 'Error') {
        return res.status(500).json({'error: ' : itemresult}); 
    }

    ewsArgs = {
        'attributes': {
            'MessageDisposition': 'SaveOnly',
        },
        'SavedItemFolderId' : {
            'DistinguishedFolderId': {
              'attributes': {
                'Id': 'calendar'
              },
              'Mailbox': {
                'EmailAddress': emailaddress
              }
            }
        },
        'Items': {
            't:CancelCalendarItem': {
                't:ReferenceItemId': {
                    'attributes': {
                        'Id': itemid,
                        'ChangeKey': changekey
                    }
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
};

exports.getCalendarAvailability = async function (req, res) {
    
    const ewsFunction = 'GetUserAvailability';
    
    const EmailAddress = req.params.emailaddress;
  
    let StartTime
    let EndTime

    try {
        StartTime = new Date(req.query.starttime);
        EndTime = new Date(req.query.endtime);
      
        StartTime = StartTime.toISOString();
        EndTime = EndTime.toISOString();
    }
    catch(err) {
        return res.status(500).json({'error: ' : err.toString()});
    }
    const ewsArgs = {
        'TimeZone': {
            'Bias': 60,
            'StandardTime' :{
                'Bias': 0,
                'Time': "01:00:00",
                'DayOrder': 4,
                'Month': 10,
                'DayOfWeek': 'Sunday',
            },
            'DaylightTime': {
                    'Bias': -60,
                    'Time': "01:00:00",
                    'DayOrder': 4,
                    'Month': 3,
                    'DayOfWeek': 'Sunday',
            }
        },
        'MailboxDataArray': {
            'MailboxData': {
                'Email': {
                    'Address': EmailAddress
                },
                'AttendeeType': 'Required',
                'ExcludeConflicts': 'false'
            }
        },
        'FreeBusyViewOptions': {
            'TimeWindow': {
                'StartTime': StartTime,
                'EndTime': EndTime
            },
            'RequestedView': 'FreeBusyMerged'
        }
    };
    
    ews.run(ewsFunction, ewsArgs)
    .then(result => {
        res.json(result);
    })
    .catch(err => {
        console.log(err.stack);
    });

};
