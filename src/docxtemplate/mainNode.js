/* eslint-disable no-param-reassign, no-console */

import path from 'path';
import fs from 'fs-extra';
import { set as timmSet } from 'timm';
import createReportBrowser from './mainBrowser';


// ==========================================
// Main
// ==========================================

// make default output filename based on template path and extension
const getDefaultOutput = (templatePath) => {
  const { dir, name, ext } = path.parse(templatePath);
  return path.join(dir, `${name}_report${ext}`);
};

const createReport = async (options) => {
  const { template, replaceImages, _probe } = options;
  // ---------------------------------------------------------
  // Load template from filesystem
  // ---------------------------------------------------------
  const buffer = await fs.readFile(template);
  const newOptions = (timmSet(options, 'template', buffer)); // returns options clone with buffer at key 'template'

  // ---------------------------------------------------------
  // Images provided as path are converted to base64 (DEPRECATED)
  // ---------------------------------------------------------
  if (replaceImages && !options.replaceImagesBase64) {
    const imgDataBase64 = {};
    const imgNames = Object.keys(replaceImages);
    for (let i = 0; i < imgNames.length; i++) {
      const imgName = imgNames[i];
      const imgPath = replaceImages[imgName];
      const imgBuf = await fs.readFile(imgPath);
      imgDataBase64[imgName] = imgBuf.toString('base64');
    }
    newOptions.replaceImagesBase64 = true;
    newOptions.replaceImages = imgDataBase64;
  }

  // ---------------------------------------------------------
  // Parse and fill template (in-memory)
  // ---------------------------------------------------------
  const report = await createReportBrowser(newOptions);
  if (_probe != null) return report;

  // ---------------------------------------------------------
  // Write the result on filesystem
  // ---------------------------------------------------------
  const output = options.output || getDefaultOutput(template);
  await fs.ensureDir(path.dirname(output));
  await fs.writeFile(output, report);
  return null;
};

// ==========================================
// Public API
// ==========================================
export default createReport;
