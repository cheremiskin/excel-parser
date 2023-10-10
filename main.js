import xlsx from 'xlsx';

// преподаватель,
// день недели,
// время,
// дисциплина,
// группа,
// аудитория

function sortObjectKeysAlphabetically(obj) {
  const sortedKeys = Object.keys(obj).sort();
  const sortedMap = new Map();
  sortedKeys.forEach(key => {
    sortedMap.set(key, obj[key]);
  });
  return sortedMap;
}

function getLettersBetween(start, end) {
  const startCharCode = start.charCodeAt(0);
  const endCharCode = end.charCodeAt(0);

  const result = [];
  for (let charCode = startCharCode; charCode < endCharCode; charCode++) {
    result.push(String.fromCharCode(charCode));
  }

  return result;
}

class ScheduleParser {
  constructor() {
    this.workbook = null;
    this.merges = new Map();
    this.groups = new Map();
    this.subjects = [];
  }

  #read(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        resolve(xlsx.read(data, { type: 'array' }));
      }
      reader.onerror = reject;

      reader.readAsArrayBuffer(file);
    });
  }

  #removeMerges() {
    this.workbook.SheetNames.forEach(sheetName => {
      const sheet = this.workbook.Sheets[sheetName];
      const schedule = xlsx.utils.sheet_to_json(sheet, {header: 'A'});

      if (sheet['!merges']) {
        sheet['!merges'].forEach(merge => {
          const {s, e} = merge;

          const cellAddress = xlsx.utils.encode_cell(s);
          if (sheet[cellAddress]) {
            const value = sheet[cellAddress].v;

            for (let row = s.r; row <= e.r; row++) {
              for (let col = s.c; col <= e.c; col++) {
                const cellAddress = xlsx.utils.encode_cell({r: row, c: col});

                // Создаем новый объект ячейки и устанавливаем значение
                const newCell = {v: value, t: 's'}; // t: 's' означает, что это строка (вы можете изменить тип, если необходимо)

                // Устанавливаем новую ячейку в лист
                sheet[cellAddress] = newCell;
              }
            }
          }
        });

        delete sheet['!merges'];
      }
    })
  }

  #parseGroups(schedule, groupsRow = 1) {
    /*
      * заполнение групп
      */
    let prevEl = null, prevI = null;
    const map = sortObjectKeysAlphabetically(schedule[groupsRow]);

    map.forEach((el, i, arr) => {
      if (prevEl !== null && prevI !== null) {
        getLettersBetween(prevI, i).forEach((key) => {
          this.groups.set(key, prevEl);
        })
      }

      prevI = i;
      prevEl = el;
    })

    this.groups.set(prevI, prevEl);
  }

  getInfo(cell) {
    const info = cell.split('\n').filter(Boolean);
    const isFullInfo = info.length >= 3;
    const isName = info.length === 1;

    return {
      info,
      isFullInfo,
      isName
    }
  }

  async parse(file) {
    this.workbook = await this.#read(file);
    this.#removeMerges();

    this.workbook.SheetNames.forEach(sheetName => {
      const sheet = this.workbook.Sheets[sheetName];
      const schedule = xlsx.utils.sheet_to_json(sheet, {header: 'A'});

      this.#parseGroups(schedule);

      const fulls = new Set();
      schedule.forEach((row, i) => {
        if (i > 1) {
          const isNumeratorRow = i + 1 !== schedule.length  && row['B'] === schedule[i + 1]['B'];

          for (let key in row) {
            if (!['A', 'B'].includes(key)) {
              const info = row[key].split('\n').filter(Boolean);
              const isFullInfo = info.length >= 3;
              const isName = info.length === 1;

              let everyWeek = false;
              let onlyNumerator = false;
              let onlyDenominator = false;

              if (isName) {
                continue;
              }

              if (isNumeratorRow) {
                const next = schedule[i + 1][key];
                if (next === undefined) {
                  onlyNumerator = true
                } else {
                  const {isName: nextIsName} = this.getInfo(next);
                  if (nextIsName) {
                    everyWeek = true;
                  } else if (next !== row[key]) {
                    onlyNumerator = true
                  } else {
                    everyWeek = true;
                  }
                }
              } else {
                const prev = schedule[i - 1][key];
                if (prev === undefined || prev !== row[key]) {
                  onlyDenominator = true
                }
              }

              if (!isFullInfo) {
                info.push(schedule[i+1][key]);
              }

              const item = {
                subject: info.slice(0, -2).join('\n'),
                teacher: info.at(-1),
                place: info.at(-2),
                group: this.groups.get(key),
                time: row['B'],
                day: row['A'],
                everyWeek,
                onlyNumerator,
                onlyDenominator
              }

              fulls.add(JSON.stringify(item));
            }
          }
        }
      })

      this.subjects.push(...Array.from(fulls).map(JSON.parse))
    })

    console.log(this.subjects)
  }
}

document.getElementById('convertButton').addEventListener('click', function() {
  const input = document.getElementById('excelFile');
  const outputDiv = document.getElementById('output');

  const file = input.files[0];
  if (!file) {
    alert('Please select an Excel file.');
    return;
  }

  const parser = new ScheduleParser();
  parser.parse(file);
});
