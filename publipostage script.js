const part1 = `
<!DOCTYPE html>
<html>
  <head>
    <style>
      table {
        font-family: arial, sans-serif;
        border-collapse: collapse;
        width: 50%;
      }
      td,
      th {
        border: 1px solid #dddddd;
        text-align: left;
        padding: 8px;
      }

      tr:nth-child(even) {
        background-color: #dddddd;
      }
      strong {
        font-size: large;
        color: red;
      }
    </style>
  </head>
  <body>
    Bonjour,<br /><br />
    Depuis quelques semaines le service de Location Decathlon est ouvert.<br />
    Grâce à vous, plusieurs centaines de commandes de Location ont déjà pu être
    remises aux clients via nos magasins.<br />
    <p>
      Vous avez déjà reçu ou allez bientôt recevoir le(s) colis retour,<br />
      <strong
        >Toutefois Suite à un problème technique, le bordereau de retour de
        certaines commandes doit être modifié</strong
      >. <br />Vous trouverez en pièce jointe le nouveau bordereau à mettre sur
      les colis avec les correspondances que voici:
      <br/>
      <i>*Notez que sur le colis la ref peut parfois (plus rarement) être le n° indiqué en Ref CLient2</i>
    </p>

    <table>
      <tr>
        <th>Ref Client(figure sur l'étiquette)</th>
        <th>Ref Client2*</th>
        <th>Nouvelle étiquette en pj</th>
        <th>Date d'expédition</th>
      </tr>
`;
const part2 = `
</table>
    <br />
    Et toujours valable: Qui puis-je contacter si j'ai un litige?<br />
    Contacte nous sur
    <a href="info-location@decathlon.com">info-location@decathlon.com</a>, c'est
    une boite que toute l'équipe reçoit et tu as l'assurance d'avoir quelqu'un
    du lundi au vendredi, de 9h à 18h. <br />
    Attention, le numéro de téléphone est réservé à vos demandes magasin
    urgentes. Merci de ne pas le communiquer aux clients!<br />
    <br />
    Un grand merci pour votre aide et une Bonne saison ! <br />
    <br />

    <img
      src="https://ci3.googleusercontent.com/proxy/MWvgSi0sXdzE9M1bhJUbDVcqKNfmeTYDXLErtStcMIac4QqQdh2Bni1IQkR0qH_q6ygnzaI7EEfrC1bKrfGuDfi4CwC3LP-GToK8COohYrCBRD1dYnA3OzMVOTAQwBIbIwbyncVdnjKbOLXp5iPIvSJAcgsRfH4UqWcZERp8mfarpacyCK9ekpcjWJfoi1hQKfUlsBEiFnezvsAmgg=s0-d-e1-ft#https://docs.google.com/uc?export=download&id=1Bj9_84_SOxIL9dtVi_CcjS5OpxmcsS-j&revid=0B11Pu44yd1GCMmgxT0poQng2V3NreUF2SXV5T2hGQ1l5T2VzPQ"
    /><br />
    Retrouve toutes les infos liées au service Location sur le blog
    <br />

    <a
      href="https://sites.google.com/decathlon.com/univers-randonnee/services/location"
      >https://sites.google.com/decathlon.com/univers-randonnee/services/location</a
    ><br />
    -- <br />
    Team location Camping/Bivouac Decathlon location-tente.decathlon.fr
  </body>
</html>
`;

function sendMail() {
  var errors = [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("missings");
  var folder = DriveApp.getFoldersByName("test").next();

  var countTotalMail = 0;
  maxRows = sheet1.getMaxRows();

  // Boucle sur chaque ligne
  for (var mainIndex = 2; mainIndex <= maxRows; mainIndex++) {
    var emailAddress = sheet1.getRange(mainIndex, 17).getValue();

    var subject = "[URGENT] Changement d'étiquette d'expédition";

    var htmlTablePart = "";
    var attached = [];

    try {
      var omsOrderIdRaw = sheet1.getRange(mainIndex, 4).getValue().toString();
      var omsOrderId = "000000000".substr(omsOrderIdRaw.length) + omsOrderIdRaw;
      htmlTablePart += `<tr><td>${omsOrderId}</td><td>${sheet1
        .getRange(mainIndex, 1)
        .getValue()}</td><td>${sheet1
        .getRange(mainIndex, 14)
        .getValue()}</td><td>${sheet1
        .getRange(mainIndex, 8)
        .getValue()
        .toString()
        .slice(0, 16)}</td></tr>`;
      // Search in google drive if file exist
      var firstFile = folder
        .getFilesByName(sheet1.getRange(mainIndex, 14).getValue())
        .next();
      attached.push(firstFile.getAs(MimeType.PDF));
      sheet1.getRange(mainIndex, 18).setValue("true");
    } catch (e) {
      errors.push({
        orderId: sheet1.getRange(mainIndex, 1).getValue(),
        label: sheet1.getRange(mainIndex, 14).getValue(),
        email: emailAddress,
      });
      sheet1.getRange(mainIndex, 18).setValue("false");
    }

    // Boucle si le mail de la ligne suivante est le même que la ligne actuelle
    for (
      var sameEmailIndex = mainIndex + 1;
      sheet1.getRange(sameEmailIndex, 17).getValue() === emailAddress;
      sameEmailIndex++
    ) {
      var omsOrderIdRaw = sheet1
        .getRange(sameEmailIndex, 4)
        .getValue()
        .toString();
      var omsOrderId = "000000000".substr(omsOrderIdRaw.length) + omsOrderIdRaw;
      htmlTablePart += `<tr><td>${omsOrderId}</td><td>${sheet1
        .getRange(sameEmailIndex, 1)
        .getValue()}</td><td>${sheet1
        .getRange(sameEmailIndex, 14)
        .getValue()}</td><td>${sheet1
        .getRange(sameEmailIndex, 8)
        .getValue()
        .toString()
        .slice(0, 16)}</td></tr>`;
      try {
        var file = folder
          .getFilesByName(sheet1.getRange(sameEmailIndex, 14).getValue())
          .next();
        attached.push(file.getAs(MimeType.PDF));
        sheet1.getRange(sameEmailIndex, 18).setValue("true");
      } catch (e) {
        errors.push({
          orderId: sheet1.getRange(sameEmailIndex, 1).getValue(),
          label: sheet1.getRange(sameEmailIndex, 14).getValue(),
          email: emailAddress,
        });
        sheet1.getRange(sameEmailIndex, 18).setValue("false");
      }
      mainIndex++;
    }

    countTotalMail++;
    const message = part1 + htmlTablePart + part2;
    if (attached.length > 0) {
      MailApp.sendEmail(emailAddress, subject, "t", {
        htmlBody: message,
        name: subject,
        attachments: attached,
      });
    }
  }
  console.log(
    `ended with ${countTotalMail} mails sent and ${errors.length} erreurs`
  );
  if (errors.length > 0) console.error(errors);
}
