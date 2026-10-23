function setupWorkspace() {
  const TREE = {
    '00_Admin': {},
    '01_Issues': {
      'Issue_001': { 'Drafts': {}, 'Final': {}, 'Assets': {} }
    },
    '02_Student_Voices': { 'Verified': {}, 'Pending': {}, 'Rejected': {} },
    '03_History_Corner': { 'Profiles': {}, 'Documents': {} },
    '04_Podcast': { 'Audio': {}, 'Transcripts': {}, 'Cover_Art': {} },
    '05_Research': { 'Sources': {}, 'Policy_Briefs': {} },
    '06_Design_Studio': {
      'Brand_Assets': {}, 'Founder_Portraits': {},
      'Canva_Exports': {}, 'Fonts_And_Palettes': {}
    },
    '07_Form_Responses': {},
    '99_Archive': {}
  };

  const CMS = {
    'Issue_Settings': ['Issue_ID','Issue_Number','Issue_Title','Publish_Date','Editor_Name','Cover_Kicker','Cover_Headline','Founder_Note','Cover_Image_URL','Status'],
    'Student_Voices': ['Voice_ID','Reference_Number','Date_Received','Submitter_Type','Display_Name','Institution','Category','Challenge','Why_It_Matters','Suggested_Solutions','Source_URL','Verification_Status','Editorial_Status','Issue_Number','Editor_Notes'],
    'History_Corner': ['History_ID','Title','Person_or_Event','Period','Institution','Achievements','Challenges','Lessons_Learned','Leadership_Model','Source_URL','Verification_Status','Editorial_Status','Issue_Number','Image_URL'],
    'Podcast': ['Episode_ID','Episode_Title','Guest','Release_Date','Summary','Audio_URL','Duration','Editorial_Status','Issue_Number'],
    'Opportunities': ['Opp_ID','Title','Organisation','Category','Deadline','Eligibility','Apply_URL','Verification_Status','Editorial_Status','Issue_Number'],
    'Submission_Log': ['Reference_Number','Timestamp','Submission_Type','Name','Email','Institution','Category','Summary','Consent_Given','Acknowledgement_Given','Verification_Status','Assigned_To','Email_Sent','Notes']
  };

  function build(parent, node) {
    Object.keys(node).forEach(function (name) {
      build(parent.createFolder(name), node[name]);
    });
  }

  const root = DriveApp.createFolder('THE MOBILIZERS DIARY');
  build(root, TREE);

  const ss = SpreadsheetApp.create('MD_CMS');
  const names = Object.keys(CMS);
  names.forEach(function (n, i) {
    const sh = i === 0 ? ss.getSheets()[0].setName(n) : ss.insertSheet(n);
    const headers = CMS[n];
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#123F3B').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setTabColor('#C59B35');
  });

  DriveApp.getFileById(ss.getId())
    .moveTo(root.getFoldersByName('00_Admin').next());

  const props = PropertiesService.getScriptProperties();
  props.setProperty('ROOT_ID', root.getId());
  props.setProperty('CMS_ID', ss.getId());

  Logger.log('ROOT: ' + root.getUrl());
  Logger.log('CMS: ' + ss.getUrl());
}


