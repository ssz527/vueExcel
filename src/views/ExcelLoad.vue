<template>
  <div>
    <div>
      <input type="file" multiple="false" id="sheetjs-input" :accept="SheetJSFT" @change="onchange" />
      <br/>
      <button type="button" id="export-table" style="visibility:hidden" @click="onexport">Export to XLSX</button>
      <br/>
      <div id="out-table"></div>
    </div>
  </div>
</template>
<script>
import XLSX from 'xlsx'
import { SheetJSFT } from '@/assets/util'

export default {
  name: 'ExcelLoad',
  data () {
    return {
      msg: 'Welcome to Your Vue.js App',
      SheetJSFT: SheetJSFT
    }
  },
  methods: {
    onchange: function (evt) {
      console.log('onchange')
      var file
      var files = evt.target.files

      if (!files || files.length === 0) return

      file = files[0]

      var reader = new FileReader()
      reader.onload = function (e) {
        console.log('reader onload')
        console.log(e)
        // pre-process data
        var binary = ''
        var bytes = new Uint8Array(e.target.result)
        var length = bytes.byteLength
        for (var i = 0; i < length; i++) {
          binary += String.fromCharCode(bytes[i])
        }

        /* read workbook */
        var wb = XLSX.read(binary, { type: 'binary' })
        console.log('wb')
        console.log(wb)

        /* grab first sheet */
        var wsname = wb.SheetNames[0]
        var ws = wb.Sheets[wsname]
        console.log('wsname')
        console.log(wsname)
        console.log('ws')
        console.log(ws)

        /* generate HTML */
        var HTML = XLSX.utils.sheet_to_html(ws)
        console.log('HTML')
        console.log(HTML)

        /* generate HTML */
        var JSON = XLSX.utils.sheet_to_json(ws)
        console.log('JSON')
        console.log(JSON)

        /* update table */
        document.getElementById('out-table').innerHTML = HTML
        /* show export button */
        document.getElementById('export-table').style.visibility = 'visible'
      }

      reader.readAsArrayBuffer(file)
    },
    onexport: function (evt) {
      /* generate workbook object from table */
      var wb = XLSX.utils.table_to_book(document.getElementById('out-table'))
      /* generate file and force a download */
      XLSX.writeFile(wb, 'sheetjs.xlsx')
    }
  }
}
</script>
