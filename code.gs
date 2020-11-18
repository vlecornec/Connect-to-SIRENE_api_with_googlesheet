function onOpen() {
  var subMenus = [{name:"Siren to Siret", functionName: "Sirentosiret"},
                  {name:"Siret Information", functionName: "Siretinfo"}
                  ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Siret/Siren", subMenus);
}
function siren1(siren) {

  var headers = {
       'Authorization': '***************************',
  };
  var options =
      {
        "method" : "get",
        "headers": headers,
        "contentType": "application/json",
        "muteHttpExceptions" : true
      };
  var result = UrlFetchApp.fetch('https://api.insee.fr/entreprises/sirene/V3/siren/' +siren + '?date=2020-11-08&champs=nicSiegeUniteLegale', options);
    var response=result.getResponseCode();
  if (response == "200" ) {
  var json = result.getContentText();
  return json}
  else{
  return response}
}
function siret1(siret) {

  var headers = {
    'Authorization': '***************************',
  };
  var options =
      {
        "method" : "get",
        "headers": headers,
        "contentType": "application/json",
        "muteHttpExceptions" : true
      };
  var result = UrlFetchApp.fetch('https://api.insee.fr/entreprises/sirene/V3/siret/' +siret + '?date=2020-11-08&champs=statutDiffusionEtablissement%2CtrancheEffectifsEtablissement%2CactivitePrincipaleRegistreMetiersEtablissement%2CetablissementSiege%2CidentifiantAssociationUniteLegale%2CtrancheEffectifsUniteLegale%2CcategorieEntreprise%2CetatAdministratifUniteLegale%2CactivitePrincipaleUniteLegale%2CcategorieJuridiqueUniteLegale%2CnomenclatureActivitePrincipaleUniteLegale%2CcaractereEmployeurUniteLegale%2CcodePostalEtablissement%2CactivitePrincipaleEtablissement', options);
    var response=result.getResponseCode();
  if (response == "200" ) {
  var json = result.getContentText();
  return json}
  else{
  return response}
}
function Siretinfo() {
    // Get the current spreadsheet, sheet, range and selected addresses
     var ss = SpreadsheetApp.getActiveSpreadsheet();
  var account_live= ss.getSheetByName('all_account_Q4');
    for (var i=7867;i<7877; i++) {
        var addresses=account_live.getRange(i,13).getValues();
        var location = siret1(addresses);
        account_live.getRange(i, 15).setValue(location);
        Utilities.sleep(1450);
    }
}
function Sirentosiret() {
    // Get the current spreadsheet, sheet, range and selected addresses
     var ss = SpreadsheetApp.getActiveSpreadsheet();
  var account_live= ss.getSheetByName('all_account_Q4');
    for (var i=1355;i<1356; i++) {
        var addresses=account_live.getRange(i,5).getValues();
        var location = siren1(addresses);
        account_live.getRange(i, 10).setValue(location);
        Utilities.sleep(1700);
    }
}
