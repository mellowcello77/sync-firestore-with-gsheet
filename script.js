// Create menu items
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Get Submissions')
//    .addItem('⏪ Export to Firestore', 'menuExport')
    .addItem('⏩ Import from Firestore', 'menuImport')
    .addToUi();
}

function menuExport() {
  main(true)
}

function menuImport() {
  main(false)
}

function main(exportData) {
  // Get the active spreadsheet and it's name for the collection name
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();
  
  // Get the first row as object properties
  // [ 'id', 'name', 'description', 'projectUrl', 'messageForParents', 'furtherResources', 'category', 'toDelete' ]
  var properties = getProperties(sheet)
  
  // Get all the data from the sheet, however many rows that may be
  // [ [ (empty), 'My Project' ,'Fun!', 'https://test.me' ,'Hello', 'some extra' ,'Test Category 1', (empty?) ], ..., ... ]
  var records = getRecords(sheet);
  
  const collectionName = 'submissions' // TODO - enter your collection name here
  
  // Values are in config file. Details on how to set it up came from here: https://github.com/grahamearley/FirestoreGoogleAppsScript
  const key = '-----BEGIN PRIVATE KEY-----\YourPrivateKeyHere\n-----END PRIVATE KEY-----\n' // TODO - enter your private key here
  const email = 'YourGoogleServiceAccountEmail@gserviceaccount.com' // TODO - enter your email here
  const projectId = 'YourProjectNameID' // TODO - Enter your project ID here
  
  var firestore = FirestoreApp.getFirestore(email, key, projectId)  
   
  if (exportData) {
    exportToFirestore(firestore, collectionName, properties, records, sheet);
  } else {
    importFromFirestore(firestore, collectionName, properties, records, sheet);
  }
}

function importFromFirestore(firestore, collectionName, properties, records, sheet) {
  // We import the data from the collection and either do an update or addition. We don't remove records that exist in the sheet but not in Firestore
  firestore.getDocuments(collectionName).forEach(function(data) {
    const splitPath = data.name.split('/');
    const documentId = splitPath[splitPath.length - 1];
    
    // For some bizarre reason the firstIndex array method wasn't working so I broke down and implemented it with a for loop
    var sheetRowIndex = -1
    for (i = 0; i < records.length; i++) {
      if (records[i][0] == documentId) {        
        sheetRowIndex = i
        break;
      }
    }

    if (sheetRowIndex < 0) {
      // Doesn't exist so we need to add it to our sheet. I'm being lazy and assume an order
      sheet.appendRow([
        documentId, 
        addField(data.fields.Premise_ID),
        addField(data.fields.Date_Submitted[0]),
        addField(data.fields.DD_Area),
        addField(data.fields.NBHD_Description),
        addField(data.fields.NBHD_ID),
        addField(data.fields.Erf),
        addField(data.fields.Erf_Description),
        addField(data.fields.LIS_Category),
        addField(data.fields.NBHD_Type),
        addField(data.fields.SR),
        addField(data.fields.On_VA18),
        addField(data.fields.PIN),
        addField(data.fields.ParentPIN),
        addField(data.fields.SR_File),
        addField(data.fields.TownName),
        addField(data.fields.Use_Description),
        addField(data.fields.LIS_Extent),
        addField(data.fields.LIS_Market_Value),
        addField(data.fields.LIS_Unit_Legal_Area),
        addField(data.fields.LIS_Unit_Nr),
        addField(data.fields.LIS_Unit_Type),
        addField(data.fields.LIS_Zone),
        addField(data.fields.Cond),
        addField(data.fields.Cp),
        addField(data.fields.Garage),
        addField(data.fields.Gf),
        addField(data.fields.NoStoreys),
        addField(data.fields.Noise),
        addField(data.fields.Note),
        addField(data.fields.Owner),
        addField(data.fields.Pool),
        addField(data.fields.Portion),
        addField(data.fields.PropertyView),
        addField(data.fields.Qual),
        addField(data.fields.Security),
        addField(data.fields.Sq),
        addField(data.fields.Street_Name),
        addField(data.fields.Street_Number),
        addField(data.fields.Submited_By_User),
        addField(data.fields.Tla1),
        addField(data.fields.Tla2),
        addField(data.fields.Tla3),
        addField(data.fields.Topo),
        addField(data.fields.Xcord),
        addField(data.fields.Ycord),
        addField(data.fields.submitLocationLat[0]),
        addField(data.fields.submitLocationLong[0]),        
        '']);
    } else {
      // We've found it so let's update the sheet
      Object.getOwnPropertyNames(data.fields).forEach(function(docProperty, docPropertyIndex) {
        // Iterate through the object properties, find the column and set the data
        
        // Important you just don't assume a precise match between the order the fields will show up from Firebase
        // and the order of your sheet. Can easily be different
        const headerColumnIndex = properties.indexOf(docProperty);
        if (headerColumnIndex >= 0) {
          // Generate the cell ID i.e. B2 or C6
          const cellId = String.fromCharCode('A'.charCodeAt() + headerColumnIndex) + (sheetRowIndex + 2);
          
          // Set the value
          sheet.setCurrentCell(sheet.getRange(cellId)).setValue(data.fields[docProperty]);
        }
      })
    }
  });
}
                       
function addField(prop) {
  // If I don't do this it will say "undefined" for fields that don't exist
  return prop ? prop : ''
}

function exportToFirestore(firestore, collectionName, properties, records, sheet) {
  var sheetRowIndex = 2
  
  records.map(function(record) {
    // record: [ (empty), 'My Project' ,'Fun!', 'https://test.me' ,'Hello', 'some extra' ,'Test Category 1' ]
    // props : [ 'id', 'name', 'description', 'projectUrl', 'messageForParents', 'furtherResources', 'category', 'toDelete' ]
    var data = {};
    properties.forEach(function(prop, i) { data[prop] = record[i]; });
    return data;
  }).forEach(function(data) {   
    const id = data.id;
    const toDelete = data.toDelete;
    delete data.id; // We don't want to store these as values so delete these properties
    delete data.toDelete;

    if (id && toDelete) {
      // Delete this row in Firestore and the sheet
      firestore.deleteDocument(collectionName + '/' + id)
      sheet.deleteRow(sheetRowIndex)
    } else if (id && !toDelete) {
      // Update this document in Firestore
      firestore.updateDocument(collectionName + '/' + id, data)

    } else if (!id && !toDelete) {
      // Create the document in Firestore and then store the ID in the sheet
      var createdRecord = firestore.createDocument(collectionName, data);
      
      // The name is the full path. The ID is at the end 'projects/<project-id>/databases/(default)/documents/<collection-name>/33ZAD2XNK98JC3HQ4hIl'
      const splitPath = createdRecord.name.split('/');
      if (splitPath.length > 0) {
        sheet.setCurrentCell(sheet.getRange('A' + sheetRowIndex)).setValue(splitPath[splitPath.length - 1]);
      }
    }
    
    if (!toDelete) {
      sheetRowIndex++;
    }
  });
}

function getProperties(sheet) {
  return sheet.getRange(1, 1, 1, 8).getValues()[0]; // [ 'id', 'name', 'description', 'projectUrl', 'messageForParents', 'furtherResources', 'category', 'toDelete' ]
}

function getRecords(sheet) {
  var data = sheet.getDataRange().getValues();
  var dataToImport = [];
  for (var i = 1; i < data.length; i++) {
    dataToImport.push(data[i]);
  }
  return dataToImport;
}