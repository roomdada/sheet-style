<template>
  <table>
    <thead>
      <th>Name</th>
      <th>Index</th>
    </thead>
    <tbody>
      <tr v-for="(row, idx) in rows" :key="idx">
        <td>{{ row.Name }}</td>
        <td>{{ row.Index }}</td>
      </tr>
    </tbody>
    <tfoot>
      <td colspan="2">
        <button @click="exportFile">Export XLSX</button>
      </td>
    </tfoot>
  </table>
</template>

<script setup>
import { ref, onMounted } from "vue";
import { read, utils } from 'xlsx';
import XLSX from 'sheetjs-style';


const rows = ref([]);

onMounted(async () => {
  /* Download from https://sheetjs.com/pres.numbers */
  const f = await fetch("https://sheetjs.com/pres.numbers");
  const ab = await f.arrayBuffer();

  /* parse workbook */
  const wb = read(ab);

  /* update data */
  rows.value = utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
});

function exportFile() {
 var workbook = XLSX.utils.book_new();

var ws = XLSX.utils.aoa_to_sheet([
    ["A1", "B1", "C1"],
    ["Da Sié Roger", "B2", "C2"],
    ["A3", "B3", "C3"]
])
ws['A2'].s = {
    font: {
        name: 'arial',
        sz: 10,
        bold: true,
        italic: true,
        border: {
            top: {style: "thin", color: {auto: 1}},
            bottom: {style: "thin", color: {auto: 1}},
            left: {style: "thin", color: {auto: 1}},
            right: {style: "thin", color: {auto: 1}}
        },
        color: {
            rgb: "2e5cdc"
        }
    },
}

ws['B2'].s = {
    font: {
        name: 'arial',
        sz: 64,
        bold: true,
        color: "green"
    },
}

XLSX.utils.book_append_sheet(workbook, ws, "SheetName");
XLSX.writeFile(workbook, 'FileName.xlsx');
}
</script>
