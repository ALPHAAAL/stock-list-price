'use client'

import ExcelJs, { CellValue } from 'exceljs';
import { useCallback, useState } from 'react';

const TABLE_HEADERS = ['股票編號', '買入價', '股數', '每股賺 ($)', '每股賺 (%)', '息率', '總回報 ($)', '總回報 (%)'];

function saveByteArray(name: string, byte: ArrayBuffer) {
  var blob = new Blob([byte], {type: "application/vnd.ms-excel"});
  var link = document.createElement('a');
  link.href = window.URL.createObjectURL(blob);
  var fileName = name;
  link.download = fileName;
  link.click();
};

type Table = Array<Array<CellValue>> | null;

export default function Home() {
  const [table, setTable] = useState<Table>(null);
  const handleFileChange = useCallback(async (e: React.FormEvent<HTMLInputElement>) => {
    const wb = new ExcelJs.Workbook();
    const excelBuffer = await (e.target as HTMLInputElement).files![0].arrayBuffer();

    await wb.xlsx.load(excelBuffer);

    const worksheet = wb.getWorksheet('Sheet1');

    let table: Table = [];

    worksheet.getRows(1, worksheet.rowCount)?.map((row, rowNumber) => {
      row.eachCell((cell) => {
          if (rowNumber !== 0) {
            const actualRowNumber = rowNumber - 1;

            table![actualRowNumber] ??= [];
            table![actualRowNumber].push(cell.value);
          }
        });
    });

    const fetchData = async (stockCode: string) => {
      const res = await fetch(`http://127.0.0.1:5001/finance-backend-6edde/asia-east2/yahooFinanceBackend/${stockCode}`);
      const result = await res.json();

      return result;
    }

    if (table) {
      const promises = table.map((val) => fetchData(val[0] as string));
      const newTable = JSON.parse(JSON.stringify(table));

      const results = await Promise.all(promises)

      results.forEach((val, i) => {
        const buyPrice = newTable[i][1];
        const lot = newTable[i][2];
        const eps = Number((val.price.regularMarketPrice - buyPrice).toPrecision(4));
        const epsp = (eps / buyPrice * 100).toPrecision(4);
        const interestRate = (val.price.regularMarketPrice * val.summaryDetail.dividendYield / buyPrice * 100).toPrecision(4);
        const totalReturn = eps * lot;
        const totalReturnPercentage = (totalReturn / (buyPrice * lot) * 100).toPrecision(4);

        newTable[i]?.push(eps, `${epsp}%`, `${interestRate}%`, totalReturn, `${totalReturnPercentage}%`);
      });

      const row = worksheet.getRow(1);

      TABLE_HEADERS.forEach((header, headerIndex) => {
        row.getCell(headerIndex + 1).value = header;
      });
      row.commit();

      newTable.forEach((row: Array<CellValue>, rowIndex: number) => {
        const excelRow = worksheet.getRow(rowIndex + 2);

        TABLE_HEADERS.forEach((_, headerIndex) => {
          const val = row[headerIndex];
          const cell = excelRow.getCell(headerIndex + 1);

          cell.value = val;

          if (typeof val === 'string' && val?.includes('%')) {
            cell.value = Number(val.slice(0, -1)) / 100;
            cell.numFmt = '0.00%';
          }

          if (typeof cell.value === 'number' && headerIndex === 7) {
            cell.style.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: {
                argb: cell.value < 0 ? 'FFFF0000' : 'FF00FF00',
              },
            }
          }
        });

        excelRow.commit();
      });

      const newWorkBook = await wb.xlsx.writeBuffer();

      saveByteArray('file.xlsx', newWorkBook);
      setTable(newTable);
    }
  }, []);

  return (
    <div>
      <input type="file" onChange={handleFileChange}/>

      {table && (
        <table className='w-full text-sm text-left text-gray-500 dark:text-gray-400'>
          <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-gray-700 dark:text-gray-400">
            <tr>
              {TABLE_HEADERS.map((val) => <th key={val} scope="col" className="px-6 py-3">{val}</th>)}
            </tr>
          </thead>
          <tbody>
            {
              table.map((row, i) => {
                return (
                  <tr key={`${row}_${i}`} className="bg-white border-b dark:bg-gray-800 dark:border-gray-700">
                    {row.map((val, i) => <td key={`${val}_${i}`} className="px-6 py-4">{val as unknown as string}</td>)}
                  </tr>
                )
              })
            }
          </tbody>
        </table>
      )}
    </div>
  )
}
