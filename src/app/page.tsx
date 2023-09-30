'use client'

import ExcelJs, { CellValue } from 'exceljs';
import { useCallback, useState } from 'react';

// const BASEPATH = 'http://127.0.0.1:5001/finance-backend-6edde/asia-east2/yahooFinanceBackend';
const BASEPATH = 'https://yahoofinancebackend-zi4qm5lvba-df.a.run.app';
const STOCK_TABLE_HEADERS = ['股票編號', '買入價', '股數', '每股賺 ($)', '每股賺 (%)', '息率', '總回報 ($)', '總回報 (%)'];
const FX_TABLE_HEADERS = ['貨幣', '買入價', '數量', '現價總數', '總回報 ($)', '總回報 (%)'];

function saveByteArray(name: string, byte: ArrayBuffer) {
  var blob = new Blob([byte], { type: "application/vnd.ms-excel" });
  var link = document.createElement('a');
  link.href = window.URL.createObjectURL(blob);
  var fileName = name;
  link.download = fileName;
  link.click();
};

type Table = {
  fx: Array<Array<CellValue>>,
  stock: Array<Array<CellValue>>,
} | null;

export default function Home() {
  const [table, setTable] = useState<Table>(null);
  const handleFileChange = useCallback(async (e: React.FormEvent<HTMLInputElement>) => {
    const wb = new ExcelJs.Workbook();
    const reader = new FileReader();

    reader.readAsArrayBuffer((e.target as HTMLInputElement).files![0]);
    reader.onload = async (evt) => {
      const excelBuffer = evt?.target?.result as ArrayBuffer;

      if (excelBuffer) {
        await wb.xlsx.load(excelBuffer);

        const stockWorksheet = wb.getWorksheet('Stock');
        const fxWorksheet = wb.getWorksheet('Currency');

        let table: Table = {
          fx: [],
          stock: [],
        };

        stockWorksheet.getRows(1, stockWorksheet.rowCount)?.map((row, rowNumber) => {
          row.eachCell((cell) => {
            if (rowNumber !== 0) {
              const actualRowNumber = rowNumber - 1;

              table!.stock[actualRowNumber] ??= [];
              table!.stock[actualRowNumber].push(cell.value);
            }
          });
        });

        fxWorksheet.getRows(1, fxWorksheet.rowCount)?.map((row, rowNumber) => {
          row.eachCell((cell) => {
            if (rowNumber !== 0) {
              const actualRowNumber = rowNumber - 1;

              table!.fx[actualRowNumber] ??= [];
              table!.fx[actualRowNumber].push(cell.value);
            }
          });
        });

        const fetchStock = async (stockCode: string) => {
          const res = await fetch(`${BASEPATH}/stock/${stockCode}`);
          const result = await res.json();

          return result;
        }

        const fetchFx = async () => {
          const res = await fetch(`${BASEPATH}/fx`);
          const result = await res.json();

          return result;
        }

        const fxRate = await fetchFx();

        if (table) {
          const stock_symbols = table.stock.map((row) => row[0]).join(',');
          const newTable = JSON.parse(JSON.stringify(table));

          const results = await fetchStock(stock_symbols as string);

          newTable.stock.forEach((val: CellValue[]) => {
            const stockSymbol = val[0] as string;
            const buyPrice = val[1] as number;
            const lot = val[2] as number;
            const eps = Number((results[stockSymbol].regularMarketPrice - buyPrice).toPrecision(4));
            const epsp = (eps / buyPrice * 100).toPrecision(4);
            const interestRate = (results[stockSymbol].regularMarketPrice * results[stockSymbol].dividendYield / buyPrice * 100).toPrecision(4);
            const totalReturn = eps * lot;
            const totalReturnPercentage = (totalReturn / (buyPrice * lot) * 100).toPrecision(4);

            val.push(eps, `${epsp}%`, `${interestRate}%`, totalReturn, `${totalReturnPercentage}%`);
          });

          const row = stockWorksheet.getRow(1);

          STOCK_TABLE_HEADERS.forEach((header, headerIndex) => {
            row.getCell(headerIndex + 1).value = header;
          });
          row.commit();

          newTable.stock.forEach((row: Array<CellValue>, rowIndex: number) => {
            const excelRow = stockWorksheet.getRow(rowIndex + 2);

            STOCK_TABLE_HEADERS.forEach((_, headerIndex) => {
              const val = row[headerIndex];
              const cell = excelRow.getCell(headerIndex + 1);

              cell.value = val;

              if (typeof val === 'string' && val?.includes('%')) {
                cell.value = Number(val.slice(0, -1)) / 100;
                cell.numFmt = '0.00%';
              }

              if (typeof cell.value === 'number' && [1, 3, 6].includes(headerIndex)) {
                cell.numFmt = '$0.0';
              }

              if (typeof cell.value === 'number' && headerIndex === 7) {
                if (cell.value !== 0) {
                  cell.font = {
                    color: {
                      argb: cell.value < 0 ? 'FFFF0000' : 'FF00FF00',
                    }
                  }
                  cell.style.fill = undefined;
                }
              }
            });

            excelRow.commit();
          });

          newTable.fx.forEach((val: CellValue[]) => {
            const currency = val[0] as string;
            const buyPrice = val[1] as number;
            const lot = val[2] as number;
            const currentPrice = fxRate[currency];
            const total = currentPrice * lot;
            const totalReturn = total - buyPrice * lot;
            const totalReturnPercentage = total === 0 ? 0 : (totalReturn / (buyPrice * lot) * 100).toPrecision(4);

            val.push(total, totalReturn, `${totalReturnPercentage}%`);
          });

          const fxRow = fxWorksheet.getRow(1);

          FX_TABLE_HEADERS.forEach((header, headerIndex) => {
            fxRow.getCell(headerIndex + 1).value = header;
          });
          fxRow.commit();

          newTable.fx.forEach((row: Array<CellValue>, rowIndex: number) => {
            const excelRow = fxWorksheet.getRow(rowIndex + 2);

            FX_TABLE_HEADERS.forEach((_, headerIndex) => {
              const val = row[headerIndex];
              const cell = excelRow.getCell(headerIndex + 1);

              cell.value = val;

              if (typeof val === 'string' && val?.includes('%')) {
                cell.value = Number(val.slice(0, -1)) / 100;
                cell.numFmt = '0.00%';
              }

              if (typeof cell.value === 'number' && [1, 3, 4].includes(headerIndex)) {
                cell.numFmt = '$00.0';
              }

              if (typeof cell.value === 'number' && headerIndex === 5) {
                if (cell.value !== 0) {
                  cell.font = {
                    color: {
                      argb: cell.value < 0 ? 'FFFF0000' : 'FF00FF00',
                    }
                  }
                  cell.style.fill = undefined;
                }
              }
            });

            excelRow.commit();
          });

          const newWorkBook = await wb.xlsx.writeBuffer();
          let fileName;
          const userAgent = window.navigator.userAgent;

          if (userAgent.match(/iPad/i) || userAgent.match(/iPhone/i)) {
            fileName = 'file';
          } else {
            fileName = 'file.xlsx';
          }
          saveByteArray(fileName, newWorkBook);
          setTable(newTable);
        }
      }
    }
  }, []);

  return (
    <div className="min-h-screen">
      <input type="file" onChange={handleFileChange} />

      {table && (
        <div>
          <p>Stock</p>
          <table className='w-full text-sm text-left text-gray-500 dark:text-gray-400'>
            <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-gray-700 dark:text-gray-400">
              <tr>
                {STOCK_TABLE_HEADERS.map((val) => <th key={val} scope="col" className="px-6 py-3">{val}</th>)}
              </tr>
            </thead>
            <tbody>
              {
                table.stock.map((row, i) => {
                  return (
                    <tr key={`${row}_${i}`} className="bg-white border-b dark:bg-gray-800 dark:border-gray-700">
                      {row.map((val, i) => <td key={`${val}_${i}`} className="px-6 py-4">{val as unknown as string}</td>)}
                    </tr>
                  )
                })
              }
            </tbody>
          </table>

          <br />
          <p>FX</p>
          <table className='w-full text-sm text-left text-gray-500 dark:text-gray-400'>
            <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-gray-700 dark:text-gray-400">
              <tr>
                {FX_TABLE_HEADERS.map((val) => <th key={val} scope="col" className="px-6 py-3">{val}</th>)}
              </tr>
            </thead>
            <tbody>
              {
                table.fx.map((row, i) => {
                  return (
                    <tr key={`${row}_${i}`} className="bg-white border-b dark:bg-gray-800 dark:border-gray-700">
                      {row.map((val, i) => <td key={`${val}_${i}`} className="px-6 py-4">{val as unknown as string}</td>)}
                    </tr>
                  )
                })
              }
            </tbody>
          </table>
        </div>
      )}

      <div className='sticky top-[100vh]'>V0.4</div>
    </div>
  )
}
