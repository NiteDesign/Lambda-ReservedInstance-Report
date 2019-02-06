
// Requirements/dependencies:
// excel4node
// fs

// Limitations:
// Only retrieves AvailabilityZone RIs

// Inlude the node_modules directory when Zipping
// and uploading to AWS Lambda

// Queries running instances and compares instances with Tag Key=ReservedInstance Value=True

// TODO:
// Include Region RIs
// Parameratization;  to Email, from Email, region, instance Tag, daysRemaining alert

const AWS = require('aws-sdk');

var ec2 = new AWS.EC2({region: 'us-east-1'});
var ses = new AWS.SES({region: 'us-east-1'});
var excel = require('excel4node');
var fs = require('fs');


//Retrieve a list of all the Lambdas
function getInstances(){
  var params = {
    Filters: [
     {
       Name: "instance-state-name",
       Values: [
         "running"
       ]
     }
   ]
  };

  ec2.describeInstances(params, function(err, data) {
    if (err) console.log(err, err.stack); // an error occurred
    else
      removeInstancesWithoutTag(data);
  });
}

function removeInstancesWithoutTag(ec2Instances){
  var reservedInstances = [];
  var nonReservedInstances = [];
  var remainingInstances = [];
  var nonTaggedInstances = [];
  var tagExists = "false"
  for(var i = 0; i < ec2Instances.Reservations.length; i++){
    tagExists = "false"
    if(ec2Instances.Reservations[i].Instances[0].Platform == "windows"){
      var platform = "windows"
    }
    else {
      var platform = "linux"
    }
    for(var t = 0; t < ec2Instances.Reservations[i].Instances[0].Tags.length; t++ ){
      if(ec2Instances.Reservations[i].Instances[0].Tags[t].Key == "ReservedInstance"){
        if(ec2Instances.Reservations[i].Instances[0].Tags[t].Value.toLowerCase() == "true"){
          reservedInstances.push(ec2Instances.Reservations[i])
          tagExists = "true"
          break;
        }else if(ec2Instances.Reservations[i].Instances[0].Tags[t].Value.toLowerCase() == "false"){
          nonReservedInstances.push(ec2Instances.Reservations[i])
          tagExists = "true"
          break;
        };
      };
    };
    if (tagExists == "false") {
        for(var t = 0; t < ec2Instances.Reservations[i].Instances[0].Tags.length; t++ ){
          if(ec2Instances.Reservations[i].Instances[0].Tags[t].Key == "Name"){
            nonTaggedInstances.push({ id: ec2Instances.Reservations[i].Instances[0].InstanceId, type: ec2Instances.Reservations[i].Instances[0].InstanceType + " (" + platform + ")", name: ec2Instances.Reservations[i].Instances[0].Tags[t].Value, region: ec2Instances.Reservations[i].Instances[0].Placement.AvailabilityZone, platform: platform});
          }
        }
    };
  };

  nonTaggedInstances.sort(compare);
  getAvailabilityZoneReservationIDs(nonTaggedInstances, reservedInstances, nonReservedInstances);
};


function getAvailabilityZoneReservationIDs(nonTaggedInstances, reservedInstances, nonReservedInstances){
  var params = {
    Filters: [
     {
       Name: "state",
       Values: [
         "active"
       ]
     },
     {
       Name: "scope",
       Values: [
         "Availability Zone"
       ]
     }
   ]
 };

 ec2.describeReservedInstances(params, function(err, data) {
   if (err) console.log(err, err.stack); // an error occurred
   else
    matchInstancetoReservation(data, nonTaggedInstances, reservedInstances, nonReservedInstances)
 });
}


function matchInstancetoReservation(reservationIDs, nonTaggedInstances, reservedInstances, nonReservedInstances){
  var matchedRIs = [];
  var instances = [];
  var unAssignedInstances = [];
  var assignedRI = "false"
  var now = new Date();
  for(var r = 0; r < reservationIDs.ReservedInstances.length; r++){
    var daysRemaining = parseInt((reservationIDs.ReservedInstances[r].End - now) / 86400000);
    matchedRIs.push([reservationIDs.ReservedInstances[r].ReservedInstancesId]);
    matchedRIs[r].push(reservationIDs.ReservedInstances[r].AvailabilityZone);
    matchedRIs[r].push(reservationIDs.ReservedInstances[r].InstanceType);
    matchedRIs[r].push(reservationIDs.ReservedInstances[r].InstanceCount);
    matchedRIs[r].push(reservationIDs.ReservedInstances[r].ProductDescription);
    matchedRIs[r].push([]);
    matchedRIs[r].push(daysRemaining);
    for(var x = 0; x < reservationIDs.ReservedInstances[r].InstanceCount; x++){
      assignedRI = "false"
      for(var i = 0; i < reservedInstances.length; i++ ){
        if(reservedInstances[i].Instances[0].Platform == "windows"){
          platform = "windows"
        }
        else {
          platform = "linux"
        }

        var n = reservationIDs.ReservedInstances[r].ProductDescription.toLowerCase().includes(platform)
        if (reservationIDs.ReservedInstances[r].InstanceType == reservedInstances[i].Instances[0].InstanceType && reservationIDs.ReservedInstances[r].AvailabilityZone == reservedInstances[i].Instances[0].Placement.AvailabilityZone && n == true) {
          //console.log("Possible MATCH")
          if(matchedRIs[r][5].length >= reservationIDs.ReservedInstances[r].InstanceCount){
            break;
          }else{
            for(var t = 0; t < reservedInstances[i].Instances[0].Tags.length; t++ ){
              if(reservedInstances[i].Instances[0].Tags[t].Key == "Name"){
                instanceName = reservedInstances[i].Instances[0].Tags[t].Value
                var instance = [reservedInstances[i].Instances[0].InstanceId, instanceName, reservationIDs.ReservedInstances[r].ReservedInstancesId ]
                matchedRIs[r][5].push(instance);
                assignedRI = "true"
                reservedInstances.splice(i,1);
                break;
              }
            }
          }

        };
       };
    };
  };

  var unAssignedInstances = [];
  for(var r = 0; r < reservedInstances.length; r++){
    if(reservedInstances[r].Instances[0].Platform == "windows"){
      var platform = "windows"
    }
    else {
      var platform = "linux"
    }
    for(var t = 0; t < reservedInstances[r].Instances[0].Tags.length; t++ ){
      if(reservedInstances[r].Instances[0].Tags[t].Key == "Name"){
        unAssignedInstances.push({ id: reservedInstances[r].Instances[0].InstanceId, type: reservedInstances[r].Instances[0].InstanceType + " (" + platform + ")", name: reservedInstances[r].Instances[0].Tags[t].Value, region: reservedInstances[r].Instances[0].Placement.AvailabilityZone});
      }
    }
  }

  unAssignedInstances.sort(compare);

  createExcel(matchedRIs, unAssignedInstances, nonReservedInstances, nonTaggedInstances)
}

function compare(a,b){
  const instanceTypeA = a.type.toUpperCase();
  const instanceTypeB = b.type.toUpperCase();

  var comparison = 0;
  if (instanceTypeA > instanceTypeB){
    comparison = 1;
  } else if (instanceTypeA < instanceTypeB){
    comparison = -1;
  }

  return comparison;
}

function createExcel(matchedRIs, unAssignedInstances, nonReservedInstances, nonTaggedInstances){
  var assignedCount = 0;
  var unUsedRICount = 0;
  var totalRICount = 0;
  var totalAssignedInstances = 0;
  //Total Assigned Instances
  for(var x = 0; x < matchedRIs.length; x++){
    totalRICount = totalRICount + matchedRIs[x][3]
    unUsedRICount = unUsedRICount + (matchedRIs[x][3] - matchedRIs[x][5].length)
    totalAssignedInstances = totalAssignedInstances + matchedRIs[x][5].length
  };


  var wb = new excel.Workbook();
  var ws = wb.addWorksheet('Sheet 1');

  var styleHeader = wb.createStyle({
    font: {
      color: '#000000',
      bold: true
    },
    border: {
      left: {
        style: "none"
      },
      right: {
        style: "none"
      },
      top: {
        style: "none"
      },
      bottom: {
        style: "medium"
      }
    }
  });

  var styleAlert = wb.createStyle({
    font: {
      color: '#FF0800',
      bold: true
    },
    border: {
      left: {
        style: "medium",
        color: "#FF0000"
      },
      right: {
        style: "medium",
        color: "#FF0000"
      },
      top: {
        style: "medium",
        color: "#FF0000"
      },
      bottom: {
        style: "medium",
        color: "#FF0000"
      }
    }
  });

  var styleInstance = wb.createStyle({
    font: {
      color: '#000000',
      bold: false
    },
    border: {
      left: {
        style: "none"
      },
      right: {
        style: "none"
      },
      top: {
        style: "none"
      },
      bottom: {
        style: "none"
      }
    }
  });

  ws.column(1).setWidth(50);

  ws.cell(1,1).string("Reserved Instance Report")
  ws.cell(1,2).string("Created: " + Date())
  ws.cell(2, 1).string("Active RIs (" + matchedRIs.length + ")")
  ws.cell(3,1).string("Total RI Instances (" + totalRICount + ")")
  ws.cell(4,1).string("Assigned Instances (" + totalAssignedInstances + ")")
  if(unUsedRICount > 0 ){
    ws.cell(5,1)
      .string("UnUSED RIs (" + unUsedRICount + ")")
      .style(styleAlert);
  }else{
    ws.cell(5,1)
      .string("UnUSED RIs (" + unUsedRICount + ")");
  }
  if(unAssignedInstances.length > 0 ){
    ws.cell(6,1)
      .string("Instances Missing RI (" + unAssignedInstances.length + ")")
      .style(styleAlert);
  }else{
    ws.cell(6,1)
      .string("Instances Missing RI (" + unAssignedInstances.length + ")");
  }

  ws.cell(7,1).string("Non RI Instances 'tag=false' (" + nonReservedInstances.length  + ")")
  ws.cell(8,1).string("Missing Tag Instances (" + nonTaggedInstances.length + ")")
  ws.cell(9,1).string("Total Running Instances (" + (totalAssignedInstances + unAssignedInstances.length + nonReservedInstances.length + nonTaggedInstances.length) + ")")

  var rowCount = 11;
  for(var x = 0; x < matchedRIs.length; x++){


    if(matchedRIs[x][3] <= matchedRIs[x][5].length){
        var style = wb.createStyle({
          font: {
            color: '#006400',
            bold: true
          },
          border: {
            left: {
              style: "none"
            },
            right: {
              style: "none"
            },
            top: {
              style: "none"
            },
            bottom: {
              style: "none"
            }
          }
        });
    }else{
      var style = wb.createStyle({
        font: {
          color: '#FF0800',
          bold: true
        },
        border: {
          left: {
            style: "medium",
            color: "#FF0000"
          },
          right: {
            style: "medium",
            color: "#FF0000"
          },
          top: {
            style: "medium",
            color: "#FF0000"
          },
          bottom: {
            style: "medium",
            color: "#FF0000"
          }
        }
      });
    }

    if (matchedRIs[x][6] <= 30){
      ws.cell(rowCount,1)
        .string("Reservation ID: " + matchedRIs[x][0])
        .style(styleAlert);
        ws.cell(rowCount,2)
          .string("Expires: " + matchedRIs[x][6] + " (days)")
          .style(styleAlert);
    }else {
      ws.cell(rowCount,1)
        .string("Reservation ID: " + matchedRIs[x][0])
        .style(style);
        ws.cell(rowCount,2)
          .string("Expires: " + matchedRIs[x][6] + " (days)")
          .style(style);
    }
    rowCount = rowCount + 1;

    ws.cell(rowCount,1)
      .string("Platform: " + matchedRIs[x][4])
      .style(styleInstance);
    rowCount = rowCount + 1;
    ws.cell(rowCount,1)
      .string("AvailabilityZone: " + matchedRIs[x][1])
      .style(styleInstance);
    rowCount = rowCount + 1;
    ws.cell(rowCount,1)
      .string("Type: " + matchedRIs[x][2])
      .style(styleInstance);
    rowCount = rowCount + 1;
    ws.cell(rowCount,1)
      .string("Instance Count: " + matchedRIs[x][3])
      .style(styleInstance);
    rowCount = rowCount + 1;
    ws.cell(rowCount,1)
      .string("Assigned Count: " + matchedRIs[x][5].length)
      .style(style);
    rowCount = rowCount + 1;
    for(var i = 0; i < matchedRIs[x][5].length; i++){

      ws.cell(rowCount,1)
        .string(matchedRIs[x][5][i][0] + " (" + matchedRIs[x][5][i][1] + ")")
        .style(styleInstance);

      rowCount = rowCount + 1;
    }
    rowCount = rowCount + 1;

  }
  rowCount = rowCount + 1;
  ws.cell(rowCount,1)
    .string("UnAssigned Instances 'tag=True' (" + unAssignedInstances.length + ")")
    .style(styleHeader);
  rowCount = rowCount + 1;
  for(var y = 0; y < unAssignedInstances.length; y++){

    ws.cell(rowCount,1).string(unAssignedInstances[y].id + " (" + unAssignedInstances[y].name + ")")
    ws.cell(rowCount,2).string(unAssignedInstances[y].type)
    ws.cell(rowCount,3).string(unAssignedInstances[y].region)
    rowCount = rowCount + 1;
  }


  rowCount = rowCount + 1;
  ws.cell(rowCount,1)
    .string("Missing 'ReservedInstance' Tag (" + nonTaggedInstances.length + ")")
    .style(styleHeader);
  rowCount = rowCount + 1;

  for(var y = 0; y < nonTaggedInstances.length; y++){
    ws.cell(rowCount,1).string(nonTaggedInstances[y].id + " (" + nonTaggedInstances[y].name + ")")
    ws.cell(rowCount,2).string(nonTaggedInstances[y].type)
    ws.cell(rowCount,3).string(nonTaggedInstances[y].region)
    rowCount = rowCount + 1;

  }

  wb.write('/tmp/ReservedInstanceReport.xlsx', function(err, data) {
    if (err) console.log(err, err.stack); // an error occurred
    else {
      sendEmail();
    }
  });
}


function sendEmail(){

  data = fs.readFileSync('/tmp/ReservedInstanceReport.xlsx');

  var ses_mail = "From: Reserved Instance Report <from@email.com>\n";
  ses_mail += "To: to@email.com\n";
  ses_mail += "Subject: AWS Reserved Instance Report\n";
  ses_mail += "MIME-Version: 1.0\n";
  ses_mail += "Content-Type: multipart/mixed; boundary=\"NextPart\"\n\n";
  ses_mail += "--NextPart\n";
  ses_mail += "Content-Type: text/html; charset=us-ascii\n\n";
  ses_mail += "<b>Reserved Instance Report</b><br/><br/>Created by Lambda: <b>Infra-RI-Report</b><br/>Triggered by Cloudwatch Rule: <b>Infra-RI-Report</b><br/>\n\n";
  ses_mail += "--NextPart\n";
  ses_mail += "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;\n";
  ses_mail += "Content-Transfer-Encoding: base64\n";
  ses_mail += "Content-Disposition: attachment; filename=\"ReservedInstanceReport.xlsx\"\n\n";
  ses_mail += data.toString('base64') + "\n\n";
  ses_mail += "--NextPart--";

  var params = {
   Destinations: [
     "To@email.com"
   ],
   RawMessage: {
    Data: ses_mail
   },
   Source: "From@email.com"
  };

  ses.sendRawEmail(params, function(err, data) {
    if (err) console.log(err, err.stack); // an error occurred
    else     console.log(data);           // successful response

  });

}

exports.handler = (event, context, callback) => {
    getInstances();
};
