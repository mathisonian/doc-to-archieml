// packages
const { load } = require('archieml');
const { google: googleApisInstance } = require('googleapis');

async function readParagraphElement(element, data) {
  // pull out the text

  const textRun = element.textRun;

  // sometimes it's not there, skip this all if so
  if (textRun) {
    // sometimes the content isn't there, and if so, make it an empty string
    
    const content = textRun.content || '';

    // step through optional text styles to check for an associated URL
    if (!textRun.textStyle) return content;
    if (!textRun.textStyle.link) return content;
    if (!textRun.textStyle.link.url) return content;

    // if we got this far there's a URL key, grab it...
    const url = textRun.textStyle.link.url;

    

    // ...but sometimes that's empty too
    if (url) {
      return `<a href="${url}">${content}</a>`;
    } else {
      return content;
    }
  } else if (element.inlineObjectElement && imageHandler) {

    const objectId = element.inlineObjectElement.inlineObjectId;

    return await imageHandler(objectId, data);

  } else {
    return '';
  }
}

async function readElements(document) {
  // prepare the text holder
  let text = '';

  // check if the body key and content key exists, and give up if not
  if (!document.body) return text;
  if (!document.body.content) return text;

  // loop through each content element in the body

  for (const element of document.body.content) {
    if (element.paragraph) {
      // get the paragraph within the element
      const paragraph = element.paragraph;

      // this is a list
      const needsBullet = paragraph.bullet != null;

      if (paragraph.elements) {
        // all values in the element
        const values = paragraph.elements;

        let idx = 0;
        for (const value of values) {
          // we only need to add a bullet to the first value, so we check
          const isFirstValue = idx === 0;
  
          // prepend an asterisk if this is a list item
          const prefix = needsBullet && isFirstValue ? '* ' : '';
  
          // concat the text
          const _text = await readParagraphElement(value, document, imageHandler)
          text += `${prefix}${_text}`;
          idx++;
        }
      }
    }
  }

  return text;
}

async function docToArchieML({
  auth,
  client,
  documentId,
  google = googleApisInstance,
  imageHandler
}) {
  // create docs client if not provided
  if (!client) {
    client = google.docs({
      version: 'v1',
      auth,
    });
  }

  // pull the data out of the doc
  const { data } = await client.documents.get({
    documentId,
  });

  // convert the doc's content to text ArchieML will understand
  const text = await readElements(data, imageHandler);

  // pass text to ArchieML and return results
  return load(text);
}

module.exports = { docToArchieML };
