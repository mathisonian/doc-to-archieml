// packages
const { load } = require('archieml');
const { google: googleApisInstance } = require('googleapis');

async function readParagraphElement(element, data, imageHandler) {
  // pull out the text

  const textRun = element.textRun;

  
  // sometimes it's not there, skip this all if so
  if (textRun) {
    // sometimes the content isn't there, and if so, make it an empty string
    // console.log(element);
    
    let content = textRun.content || '';

    // step through optional text styles to check for an associated URL
    if (!textRun.textStyle) return content;

    // console.log(textRun);
    if (textRun.textStyle.italic) {
      content = `<em>${content}</em>`
    }
    if (textRun.textStyle.bold) {
      content = `<b>${content}</b>`
    }

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
  } else if (element.horizontalRule) {
    return `<hr />`;
  } else {
    return '';
  }
}

async function readElements(document, imageHandler) {
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

        const taggedText = (text) => {
          if (paragraph.paragraphStyle && paragraph.paragraphStyle.namedStyleType.includes('HEADING')) {
            const headingLevel = paragraph.paragraphStyle.namedStyleType.replace('HEADING_', '');
            return `<h${headingLevel}>${text.trim()}</h${headingLevel}>\n`;
          }
          return text;
        };
        let elementText = '';
        for (const value of values) {
          // we only need to add a bullet to the first value, so we check
          const isFirstValue = idx === 0;
  
          // prepend an asterisk if this is a list item
          const prefix = needsBullet && isFirstValue ? '* ' : '';
  
          // concat the text
          const _text = await readParagraphElement(value, document, imageHandler)
          elementText += `${prefix}${_text}`;
          idx++;
        }
        text += taggedText(elementText);
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

  let text = await readElements(data, imageHandler);

  const refs = (text.match(/{ref}(.*?){\/ref}/gims) || []).map(function(
    val,
    i
  ) {
    text = text.replace(val, `<ref id="${i}" />`);
    return val.replace(/\{\/?ref\}/g, '');
  });

  const parsed = load(text);
  
  parsed.refs = refs;

  // Parse lists and include lowercase vals
  parsed.body = parsed.body.reduce((memo, d, i) => {
    
    Object.keys(d).forEach((k) => {
      d[k.toLowerCase()] = d[k];
    })

    if (d.type === 'text' && d.value.startsWith('* ')) {
      if (memo.isInList) {
        memo.body[memo.body.length - 1].value.push(d.value.replace('* ', '').trim());
      } else {
        memo.isInList = true;
        memo.body.push({
          type: 'list',
          value: [d.value.replace('* ', '').trim()]
        });
      }
    } else {
      if (memo.isInList) {
        memo.isInList = false;
      }
      memo.body.push(d);
    }
    return memo;
  }, { 
    isInList: false,
    body: [],
  }).body;

  // pass text to ArchieML and return results
  return parsed;
}

module.exports = { docToArchieML };
