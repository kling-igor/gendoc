/* eslint-disable no-param-reassign, no-console */

import JSZip from 'jszip';
import { parseXml, buildXml } from './xml';
import preprocessTemplate from './preprocessTemplate';
import { extractQuery, produceJsReport } from './processTemplate';
import { addChild, newNonTextNode } from './reportUtils';

const DEFAULT_CMD_DELIMITER = '+++';
const DEFAULT_LITERAL_XML_DELIMITER = '||';

// ==========================================
// Main
// ==========================================
const createReport = async (options) => {
  console.log('Report options:', { attach: options });
  const { template, data, queryVars, replaceImages, _probe } = options;

  const templatePath = 'word';
  const literalXmlDelimiter = options.literalXmlDelimiter || DEFAULT_LITERAL_XML_DELIMITER;
  const createOptions = {
    cmdDelimiter: options.cmdDelimiter || DEFAULT_CMD_DELIMITER,
    literalXmlDelimiter,
    processLineBreaks:
      options.processLineBreaks != null ? options.processLineBreaks : true,
    noSandbox: options.noSandbox || false,
    additionalJsContext: options.additionalJsContext || {},
  };
  const xmlOptions = { literalXmlDelimiter };

  // ---------------------------------------------------------
  // Unzip
  // ---------------------------------------------------------
  console.log('Unzipping...');
  const zip = await JSZip.loadAsync(template);

  // ---------------------------------------------------------
  // Read the 'document.xml' file (the template) and parse it
  // ---------------------------------------------------------
  console.log('Reading template...');
  const templateXml = await zip.file(`${templatePath}/document.xml`).async('text')//zipGetText(zip, `${templatePath}/document.xml`);
  console.log(`Template file length: ${templateXml.length}`);
  console.log('Parsing XML...');
  const tic = new Date().getTime();
  const parseResult = await parseXml(templateXml);
  const jsTemplate = parseResult;
  const tac = new Date().getTime();
  // DEBUG &&
  //   console.log(`File parsed in ${tac - tic} ms`, {
  //     attach: jsTemplate,
  //     attachLevel: 'trace',
  //   });

  // ---------------------------------------------------------
  // Fetch the data that will fill in the template
  // ---------------------------------------------------------
  let queryResult = null;
  if (typeof data === 'function') {
    console.log('Looking for the query in the template...');
    const query = await extractQuery(jsTemplate, createOptions);
    console.log(`Query: ${query || 'no query found'}`);
    queryResult = await data(query, queryVars);
  } else {
    queryResult = data;
  }

  // ---------------------------------------------------------
  // Generate the report
  // ---------------------------------------------------------
  console.log('Before preprocessing...', {
    attach: jsTemplate,
    attachLevel: 'debug',
    ignoreKeys: ['_parent', '_fTextNode', '_attrs'],
  });

  const finalTemplate = preprocessTemplate(jsTemplate, createOptions);

  console.log('Generating report...', {
    attach: finalTemplate,
    attachLevel: 'debug',
    ignoreKeys: ['_parent', '_fTextNode', '_attrs'],
  });
  const { report, images } = await produceJsReport(queryResult, finalTemplate, createOptions);

  if (_probe === 'JS') return report;

  // ---------------------------------------------------------
  // Build output XML and write it to disk
  // ---------------------------------------------------------
  // DEBUG &&
  //   console.log('Report', {
  //     attach: report,
  //     attachLevel: 'debug',
  //     ignoreKeys: ['_parent', '_fTextNode', '_attrs'],
  //   });
  console.log('Converting report to XML...');
  const reportXml = buildXml(report, xmlOptions);
  if (_probe === 'XML') return reportXml;
  console.log('Writing report...');
  zip.file(`${templatePath}/document.xml`, reportXml)// zipSetText(zip, `${templatePath}/document.xml`, reportXml);

  // ---------------------------------------------------------
  // Add images
  // ---------------------------------------------------------
  console.log('Processing images...');
  const imageIds = Object.keys(images);
  if (imageIds.length) {
    console.log('Completing document.xml.rels...');
    const relsPath = `${templatePath}/_rels/document.xml.rels`;
    const relsXml = await zip.file(relsPath).async('text')//zipGetText(zip, relsPath);
    const rels = await parseXml(relsXml);
    for (let i = 0; i < imageIds.length; i++) {
      const imageId = imageIds[i];
      const { extension, data: imgData } = images[imageId];
      const imgName = `template_image${i + 1}${extension}`;
      console.log(`Writing image ${imageId} (${imgName})...`);
      const imgPath = `${templatePath}/media/${imgName}`;
      if (typeof imgData === 'string') {
        await zip.file(imgPath, imgData, { base64: true }) //zipSetBase64(zip, imgPath, imgData);
      } else {
        await zip.file(imgPath, imgData, { binary: true }) //zipSetBinary(zip, imgPath, imgData);
      }
      addChild(
        rels,
        newNonTextNode('Relationship', {
          Id: imageId,
          Type:
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          Target: `media/${imgName}`,
        })
      );
    }
    const finalRelsXml = buildXml(rels, xmlOptions);
    zip.file(relsPath, finalRelsXml) //zipSetText(zip, relsPath, finalRelsXml);

    // Process [Content_Types].xml
    console.log('Completing [Content_Types].xml...');
    const contentTypesPath = '[Content_Types].xml';
    const contentTypesXml = await zip.file(contentTypesPath).async('text')//zipGetText(zip, contentTypesPath);
    const contentTypes = await parseXml(contentTypesXml);
    console.log('Content types', { attach: contentTypes });
    const ensureContentType = (extension, contentType) => {
      const children = contentTypes._children;
      if (
        children.filter(o => !o._fTextNode && o._attrs.Extension === extension)
          .length
      ) {
        return;
      }
      addChild(
        contentTypes,
        newNonTextNode('Default', {
          Extension: extension,
          ContentType: contentType,
        })
      );
    };
    ensureContentType('png', 'image/png');
    ensureContentType('jpg', 'image/jpeg');
    ensureContentType('jpeg', 'image/jpeg');
    ensureContentType('gif', 'image/gif');
    const finalContentTypesXml = buildXml(contentTypes, xmlOptions);
    zip.file(contentTypesPath, finalContentTypesXml)//zipSetText(zip, contentTypesPath, finalContentTypesXml);
  }

  // ---------------------------------------------------------
  // Replace images
  // ---------------------------------------------------------
  if (replaceImages) {
    console.log('Replacing images...');
    if (options.replaceImagesBase64) {
      const mediaPath = `${templatePath}/media`;
      const imgNames = Object.keys(replaceImages);
      for (let i = 0; i < imgNames.length; i++) {
        const imgName = imgNames[i];
        const imgPath = `${mediaPath}/${imgName}`;

        if (zip.file(`${imgPath}`) == null)
          // if (!zipExists(zip, `${imgPath}`)) {
          console.warn(
            `Image ${imgName} cannot be replaced: destination does not exist`
          );
        continue;
      }
      const imgData = replaceImages[imgName];
      console.log(`Replacing ${imgName} with <base64 buffer>...`);
      await zip.file(filename, data, { base64: true })//zipSetBase64(zip, imgPath, imgData);
    }
  } else {
    console.warn(
      'Unsupported format (path): images can only be replaced in base64 mode'
    );
  }

  // ---------------------------------------------------------
  // Process all other XML files (they may contain headers, etc.)
  // ---------------------------------------------------------
  const files = [];
  zip.forEach(async filePath => {
    const regex = new RegExp(`${templatePath}\\/[^\\/]+\\.xml`);
    if (regex.test(filePath) && filePath !== `${templatePath}/document.xml`) {
      files.push(filePath);
    }
  });

  for (let i = 0; i < files.length; i++) {
    const filePath = files[i];
    console.log(`Processing ${filePath}...`);
    const raw = await zip.file(filePath).async('text')//zipGetText(zip, filePath);
    const js0 = await parseXml(raw);
    const js = preprocessTemplate(js0, createOptions);
    const { report: report2 } = await produceJsReport(queryResult, js, createOptions);
    const xml = buildXml(report2, xmlOptions);
    zip.file(filePath, xml)//zipSetText(zip, filePath, xml);
  }

  // ---------------------------------------------------------
  // Zip the results
  // ---------------------------------------------------------
  console.log('Zipping...');
  const output = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 1 } })//zipSave(zip);
  return output;
};

// ==========================================
// Public API
// ==========================================
export default createReport;
