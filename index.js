const docusign = require('docusign-esign');
const fs = require('fs');
const path = require('path');
const config = require('config')
const promisify = require('util').promisify;
const notifier = require('mail-notifier');
const HTMLParser = require('node-html-parser');

let data = {}

const imap = config.get('imap')
const { accessToken, accountId } = config.get('docusignConfig')

fs.readFileAsync = function(filename) {
  return new Promise(function(resolve, reject) {
      fs.readFile(filename, function(err, data){
          if (err) 
              reject(err); 
          else
              resolve(data);
      });
  });
};

notifier({...imap})
  .on('mail', async mail => {
    console.log("You got mail")
    
    let msg = mail.html;
    let start = msg.indexOf('<table');
    let end = msg.indexOf('</table>') + 8;
    msg = msg.substring(start, end);
  
    const root = HTMLParser.parse(msg);

    let rows = root.querySelectorAll(`tr`);

    data.name =      rows[1].childNodes[1].rawText.trim();
    data.email =     rows[2].childNodes[1].rawText.trim();
    data.nowater =   rows[3].childNodes[1].rawText.trim();
    data.limited =   rows[4].childNodes[1].rawText.trim();
    data.fullwater = rows[5].childNodes[1].rawText.trim();
    data.finance =   rows[6].childNodes[1].rawText.trim();
    data.allrisk =   rows[7].childNodes[1].rawText.trim();
    data.bass =      rows[8].childNodes[1].rawText.trim();

    console.log('Data received:', data)

    let docs = []
    if (data.allrisk === "1") {
      const content = await fs.readFileAsync('./templates/Allrisk.json')
      const jsonContent = updateDocumentIds(JSON.parse(content), docs.length + 1)
      docs.push(jsonContent)
    }
    if (data.bass === "1") {
      const content = await fs.readFileAsync('./templates/Bass.json')
      const jsonContent = updateDocumentIds(JSON.parse(content), docs.length + 1)
      docs.push(jsonContent)
    }
    if (data.finance === "1") {
      const content = await fs.readFileAsync('./templates/Finance.json')
      const jsonContent = updateDocumentIds(JSON.parse(content), docs.length + 1)
      docs.push(jsonContent)
    }
    if (data.fullwater === "0") {
      if (data.nowater === "1") {
        const content = await fs.readFileAsync('./templates/No_Water.json')
        const jsonContent = updateDocumentIds(JSON.parse(content), docs.length + 1)
        docs.push(jsonContent)
      } else {
        const content = await fs.readFileAsync('./templates/Limited.json')
        const jsonContent = updateDocumentIds(JSON.parse(content), docs.length + 1)
        docs.push(jsonContent)
      }
    }
    const content = await fs.readFileAsync('./templates/Flood_Rejection.json')
    const jsonContent = updateDocumentIds(JSON.parse(content), docs.length + 1)
    docs.push(jsonContent)

    // Recipient Information:

    const signerName =  data.name;

    const signerEmail = data.email;

    /**

      *  The envelope is sent to the provided email address. 

      *  One signHere tab is added.

      *  The document path supplied is relative to the working directory 

      */

    const apiClient = new docusign.ApiClient();

    apiClient.setBasePath('https://demo.docusign.net/restapi');

    apiClient.addDefaultHeader('Authorization', 'Bearer ' + accessToken);

    // Set the DocuSign SDK components to use the apiClient object

    docusign.Configuration.default.setDefaultApiClient(apiClient);



    // Create the envelope request

    // Start with the request object

    const envDef = new docusign.EnvelopeDefinition();

    //Set the Email Subject line and email message

    envDef.emailSubject = data.subject || 'Please review and sign the documents.';

    envDef.emailBlurb = data.subject || 'Please review and sign the documents.'


    // Create the document request object

    ///////////////////////////////////////////////////////////////////////////////////

    let documents = []
    docs.forEach(doc => {
      doc.documents.forEach(document => {
        documents.push(document)
      })
    })
    envDef.documents = documents


    // Add the recipients object to the envelope definition.

    // It includes an array of the signer objects. 

    let signer1
    docs[0].recipients.signers[0].name = signerName;
    docs[0].recipients.signers[0].email = signerEmail;
    signer1 = docs[0].recipients.signers[0]

    let signHereTabs = []
    docs.forEach(doc => {
      if (doc.recipients.signers[0].tabs.signHereTabs)
        doc.recipients.signers[0].tabs.signHereTabs.forEach(signHereTab => {
          signHereTabs.push(signHereTab)
        })
    })
    let dateSignedTabs = []
    docs.forEach(doc => {
      if (doc.recipients.signers[0].tabs.dateSignedTabs)
        doc.recipients.signers[0].tabs.dateSignedTabs.forEach(dateSignedTab => {
          dateSignedTabs.push(dateSignedTab)
        })
    })
    let fullNameTabs = []
    docs.forEach(doc => {
      if (doc.recipients.signers[0].tabs.fullNameTabs)
        doc.recipients.signers[0].tabs.fullNameTabs.forEach(fullNameTab => {
          fullNameTabs.push(fullNameTab)
        })
    })
    let initialHereTabs = []
    docs.forEach(doc => {
      if (doc.recipients.signers[0].tabs.initialHereTabs)
        doc.recipients.signers[0].tabs.initialHereTabs.forEach(initialHereTab => {
          initialHereTabs.push(initialHereTab)
        })
    })

    signer1.tabs.signHereTabs = signHereTabs
    signer1.tabs.dateSignedTabs = dateSignedTabs
    signer1.tabs.fullNameTabs = fullNameTabs
    signer1.tabs.initialHereTabs = initialHereTabs
    
    let recipients = docusign.Recipients.constructFromObject({
      signers: [signer1]
    });
    envDef.recipients = recipients;

    // Set the Envelope status. For drafts, use 'created' To send the envelope right away, use 'sent'

    envDef.status = 'sent';

  /////////////////////////////////////////////////////////////////////
    
    // Send the envelope

    // The SDK operations are asynchronous, and take callback functions.

    // However we'd pefer to use promises.

    // So we create a promise version of the SDK's createEnvelope method.

    let envelopesApi = new docusign.EnvelopesApi()

        // createEnvelopePromise returns a promise with the results:

      , createEnvelopePromise = promisify(envelopesApi.createEnvelope).bind(envelopesApi)

      , results

      ;



    try {

      results = await createEnvelopePromise(accountId, {'envelopeDefinition': envDef})

    } catch  (e) {

      let body = e.response && e.response.body;

      if (body) {

        // DocuSign API exception

        console.log(`API problem - Status code ${e.response.status}`);
        console.log(`Error message: ${JSON.stringify(body, null, 4)}`)

      } else {

        // Not a DocuSign exception

        throw e;

      }

    }

    // Envelope has been created:

    if (results) {

      console.log(`Signer: ${signerName} <${signerEmail}>`);
      console.log(`Results: ${JSON.stringify(results, null, 4)}`)

    }
  })
  .start()

  const updateDocumentIds = (jsonContent, documentId) => {
    jsonContent.documents[0].documentId = documentId

    let count = 0
    count = jsonContent.recipients.signers[0].tabs.dateSignedTabs ? jsonContent.recipients.signers[0].tabs.dateSignedTabs.length : 0
    for (let i = 0; i < count; i++) {
      jsonContent.recipients.signers[0].tabs.dateSignedTabs[i].documentId = documentId
    }

    count = jsonContent.recipients.signers[0].tabs.fullNameTabs ? jsonContent.recipients.signers[0].tabs.fullNameTabs.length : 0
    for (let i = 0; i < count; i++) {
      jsonContent.recipients.signers[0].tabs.fullNameTabs[i].documentId = documentId
    }

    count = jsonContent.recipients.signers[0].tabs.initialHereTabs ? jsonContent.recipients.signers[0].tabs.initialHereTabs.length : 0
    for (let i = 0; i < count; i++) {
      jsonContent.recipients.signers[0].tabs.initialHereTabs[i].documentId = documentId
    }

    count = jsonContent.recipients.signers[0].tabs.signHereTabs ? jsonContent.recipients.signers[0].tabs.signHereTabs.length : 0
    for (let i = 0; i < count; i++) {
      jsonContent.recipients.signers[0].tabs.signHereTabs[i].documentId = documentId
    }

    return jsonContent
  }