/* eslint-disable no-param-reassign, no-console */

// import path from 'path'
// import fs from 'fs-extra'
import JSZip from 'jszip'
import { set as timmSet } from 'timm'
import fse from "fs-extra"
// import { encode, decode } from 'base64-arraybuffer'
import { parseXml, buildXml } from './xml'
import preprocessTemplate from './preprocessTemplate'
import TemplateProcessor from './TemplateProcessor'
import { addChild, newNonTextNode } from './reportUtils'

const DEFAULT_CMD_DELIMITER = '+++'
const DEFAULT_LITERAL_XML_DELIMITER = '||'

// import createReportBrowser from './mainBrowser'

class DocxTemplate {
  constructor(onReadImage) {
    this.templateProcessor = new TemplateProcessor(onReadImage)
  }

  // getDefaultOutput = (templatePath) => {
  //   const { dir, name, ext } = path.parse(templatePath);
  //   return path.join(dir, `${name}_report${ext}`);
  // };

  async createReport(options) {

    const { template, reportPath } = options

    // ---------------------------------------------------------
    // Load template from filesystem
    // ---------------------------------------------------------

    // let base64Buffer
    // try {
    //   base64Buffer = await fse.readFile(template)
    // } catch (e) {
    //   return false
    // }

    const buffer = await fse.readFile(template)


    // convert base64 to ArrayBuffer
    // const buffer = decode(base64Buffer)

    const newOptions = timmSet(options, 'template', buffer) // returns options clone with buffer at key 'template'

    // ---------------------------------------------------------
    // Images provided as path are converted to base64 (DEPRECATED)
    // ---------------------------------------------------------
    // if (replaceImages && !options.replaceImagesBase64) {
    //   const imgDataBase64 = {};
    //   const imgNames = Object.keys(replaceImages);
    //   for (let i = 0; i < imgNames.length; i++) {
    //     const imgName = imgNames[i];
    //     const imgPath = replaceImages[imgName];
    //     const imgBuf = await fs.readFile(imgPath);
    //     imgDataBase64[imgName] = imgBuf.toString('base64');
    //   }
    //   newOptions.replaceImagesBase64 = true;
    //   newOptions.replaceImages = imgDataBase64;
    // }

    // ---------------------------------------------------------
    // Parse and fill template (in-memory)
    // ---------------------------------------------------------
    const report = await this._doJob(newOptions)
    // if (_probe != null) return report // ??

    // ---------------------------------------------------------
    // Write the result on filesystem
    // ---------------------------------------------------------
    // const { output } = options // || getDefaultOutput(template);

    // const fileInfo = {
    //   name: `Report_${Date.now()}.docx`,
    //   type: 'docx'
    // }

    return await fse.writeFile(reportPath, report)
  }

  async _doJob(options) {

    const { template, data, queryVars, replaceImages, _probe } = options

    const templatePath = 'word';
    const literalXmlDelimiter = options.literalXmlDelimiter || DEFAULT_LITERAL_XML_DELIMITER
    const createOptions = {
      cmdDelimiter: options.cmdDelimiter || DEFAULT_CMD_DELIMITER,
      literalXmlDelimiter,
      processLineBreaks: options.processLineBreaks != null ? options.processLineBreaks : true,
      noSandbox: options.noSandbox || false,
      additionalJsContext: options.additionalJsContext || {}
    }
    const xmlOptions = { literalXmlDelimiter }

    // ---------------------------------------------------------
    // Unzip
    // ---------------------------------------------------------
    const zip = await JSZip.loadAsync(template)

    // ---------------------------------------------------------
    // Read the 'document.xml' file (the template) and parse it
    // ---------------------------------------------------------
    const templateXml = await zip.file(`${templatePath}/document.xml`).async('text')//zipGetText(zip, `${templatePath}/document.xml`);
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
    let queryResult = null
    if (typeof data === 'function') {
      const query = await this.templateProcessor.extractQuery(jsTemplate, createOptions);
      queryResult = await data(query, queryVars)
    } else {
      queryResult = data
    }

    // ---------------------------------------------------------
    // Generate the report
    // ---------------------------------------------------------

    const finalTemplate = preprocessTemplate(jsTemplate, createOptions);

    const { report, images } = await this.templateProcessor.produceJsReport(queryResult, finalTemplate, createOptions);

    if (_probe === 'JS') return report

    // ---------------------------------------------------------
    // Build output XML and write it to disk
    // ---------------------------------------------------------
    // DEBUG &&
    //   console.log('Report', {
    //     attach: report,
    //     attachLevel: 'debug',
    //     ignoreKeys: ['_parent', '_fTextNode', '_attrs'],
    //   });
    const reportXml = buildXml(report, xmlOptions);
    if (_probe === 'XML') return reportXml;
    zip.file(`${templatePath}/document.xml`, reportXml)// zipSetText(zip, `${templatePath}/document.xml`, reportXml);

    // ---------------------------------------------------------
    // Add images
    // ---------------------------------------------------------
    const imageIds = Object.keys(images);
    if (imageIds.length) {
      const relsPath = `${templatePath}/_rels/document.xml.rels`;
      const relsXml = await zip.file(relsPath).async('text')//zipGetText(zip, relsPath);
      const rels = await parseXml(relsXml);
      for (let i = 0; i < imageIds.length; i++) {
        const imageId = imageIds[i];
        const { extension, data: imgData } = images[imageId];
        const imgName = `template_image${i + 1}${extension}`;
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
      const contentTypesPath = '[Content_Types].xml';
      const contentTypesXml = await zip.file(contentTypesPath).async('text')//zipGetText(zip, contentTypesPath);
      const contentTypes = await parseXml(contentTypesXml);
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
            ContentType: contentType
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
      if (options.replaceImagesBase64) {
        const mediaPath = `${templatePath}/media`;
        const imgNames = Object.keys(replaceImages);
        for (let i = 0; i < imgNames.length; i++) {
          const imgName = imgNames[i];
          const imgPath = `${mediaPath}/${imgName}`;

          if (zip.file(`${imgPath}`) == null)
            // if (!zipExists(zip, `${imgPath}`)) {
            continue
        }
        const imgData = replaceImages[imgName];
        await zip.file(filename, data, { base64: true })//zipSetBase64(zip, imgPath, imgData);
      }
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
      const raw = await zip.file(filePath).async('text')//zipGetText(zip, filePath);
      const js0 = await parseXml(raw);
      const js = preprocessTemplate(js0, createOptions);
      const { report: report2 } = await this.templateProcessor.produceJsReport(queryResult, js, createOptions);
      const xml = buildXml(report2, xmlOptions);
      zip.file(filePath, xml)//zipSetText(zip, filePath, xml);
    }

    // ---------------------------------------------------------
    // Zip the results
    // ---------------------------------------------------------
    const output = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 1 } })//zipSave(zip);
    return output
  }
}

export default DocxTemplate
