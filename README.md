GenDoc
======

A small utility to genereate .docx and .xlsx by specified template and mapping data.

## Installation

  npm install gendoc.tar

## Usage

```js
import gendoc from 'gendoc'

const PATH_TO_TEMPLATE_FILE = ... // path to template file
const MAPPING_DATA = ...          // object with placeholder names containing data
const PATH_TO_REPORT_FILE = ...   // path to result file
const PLACEHOLDER_DELIMETER = '#' // default value - may be ommited

gendoc(PATH_TO_TEMPLATE_FILE, PATH_TO_REPORT_FILE, MAPPING_DATA, PLACEHOLDER_DELIMETER)
.then(()=>{
  console.log('DONE')
})
.catch(e => {
  console.error(e)
})

```

## Release History

* 1.0.0 Initial release