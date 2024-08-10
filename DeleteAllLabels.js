// Script to delete all labels

function deleteAllLabels() {
  // Get all labels in Gmail
  var allLabels = GmailApp.getUserLabels();
  
  // Loop through each label and delete it
  allLabels.forEach(function(label) {
    Logger.log('Deleting label: ' + label.getName());
    label.deleteLabel();
  });
  
  Logger.log('All labels have been deleted.');
}
