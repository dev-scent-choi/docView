import { XMLParser } from 'fast-xml-parser';

const xml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:body><w:p><w:r><w:t xml:space="preserve">First paragraph</w:t></w:r></w:p><w:sectPr/></w:body></w:document>';

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  textNodeName: '#text',
  trimValues: false,
});

console.log(JSON.stringify(parser.parse(xml), null, 2));
