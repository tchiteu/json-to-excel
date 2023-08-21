import XLSX from 'sheetjs-style';

function colorCell (worksheet, origin, color) {
  const cellAdress = XLSX.utils.encode_cell(origin);
  const borderStyle = { style: 'thin', color: { rgb: '999999' } };
  const border = {
    top: borderStyle,
    right: borderStyle,
    bottom: borderStyle,
    left: borderStyle,
  };

  worksheet[cellAdress].s = {
    fill: {
      fgColor: { rgb: color },
    },
    border,
  };
}

const rows = [
  {
    actor: 'Regina King',
    age: 25,
    awards: 48,
    movies: 23,
  },
  {
    actor: 'Gary Anthony',
    age: 25,
    awards: 23,
    movies: 12,
  },
  {
    actor: 'Jill Talley',
    age: 25,
    awards: 7,
    movies: 18,
  },
  {
    actor: 'Kevin Richardson',
    age: 25,
    awards: 1,
    movies: 4,
  },
  
];

const cellsAmount = Object.keys(rows[0]).length - 1;

const customCells = [
  {
    value: ['Actors Awards'],
    start: { r: 0, c: 0 },
    end: { r: 0, c: cellsAmount },
    color: '6e54aa',
  },
];

const cellsColors = [
  {
    color: '7860b4',
    start: { r: 1, c: 0 },
    end: { r: 1, c: cellsAmount },
  },
  {
    color: 'ccbbec',
    start: { r: 2, c: 0 },
    end: { r: rows.length, c: cellsAmount },
  },
];

const worksheet = XLSX.utils.json_to_sheet(rows, { start: 'A2' });
worksheet['!merges'] = [];

// apply custom headers
XLSX.utils.sheet_add_aoa(worksheet, [['Actor', 'Age', 'Awards', 'Movies']], { origin: 'A2' });

// apply custom cells
customCells.forEach(({ value, start, end, color }) => {
  XLSX.utils.sheet_add_aoa(worksheet, [[value]], { start });

  colorCell(worksheet, start, color);

  if (end) {
    worksheet["!merges"].push({ s: start, e: end });
  }
});

// apply cells styles
cellsColors.forEach(({ color, start, end }) => {
  const rowsLimit = end?.r || start.r;
  const cellsLimit = end?.c || start.c;

  // rows
  for (let rowIndex = start.r; rowIndex <= rowsLimit; rowIndex++) {
    // cells
    for (let cellIndex = start.c; cellIndex <= cellsLimit; cellIndex++) {
      const start = { r: rowIndex, c: cellIndex };
      colorCell(worksheet, start, color);
    }
  }
});

// create an XLSX file and try to save to Awards.xlsx
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Awards');
XLSX.writeFile(workbook, 'Awards.xlsx', { compression: true });
