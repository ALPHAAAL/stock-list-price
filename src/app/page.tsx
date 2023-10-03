'use client'

import ExcelJs, { CellFormulaValue, CellValue } from 'exceljs';
import { useCallback, useState } from 'react';
import { saveAs } from 'file-saver';

// const BASEPATH = 'http://127.0.0.1:5001/finance-backend-6edde/asia-east2/yahooFinanceBackend';
const BASEPATH = 'https://yahoofinancebackend-zi4qm5lvba-df.a.run.app';
const STOCK_TABLE_HEADERS = ['股票編號', '買入價', '股數', '現價', '每股賺 ($)', '每股賺 (%)', '息率', '總回報 ($)', '總回報 (%)'];
const FX_TABLE_HEADERS = ['貨幣', '買入價', '數量', '現價', '現價總數', '總回報 ($)', '總回報 (%)'];

function saveByteArray(name: string, byte: ArrayBuffer) {
  const blob = new Blob([byte], { type: "application/vnd.ms-excel" });

  saveAs(blob, name);
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

        let stockWorksheet = wb.getWorksheet('Stock');
        let fxWorksheet = wb.getWorksheet('Currency');

        let table: Table = {
          fx: [],
          stock: [],
        };

        stockWorksheet.getRows(1, stockWorksheet.rowCount)?.map((row, rowNumber) => {
          row.eachCell((cell, colIndex) => {
            if (rowNumber !== 0 && colIndex <= 3) {
              const actualRowNumber = rowNumber - 1;

              table!.stock[actualRowNumber] ??= [];
              table!.stock[actualRowNumber].push(cell.value);
            }
          });
        });

        fxWorksheet.getRows(1, fxWorksheet.rowCount)?.map((row, rowNumber) => {
          row.eachCell((cell, colIndex) => {
            if (rowNumber !== 0 && colIndex <= 3) {
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
            val[0] = (typeof val[0] === 'object' ? (val[0] as unknown as CellFormulaValue).result : val[0]) as string;
            val[1] = (typeof val[1] === 'object' ? (val[1] as unknown as CellFormulaValue).result : val[1]) as number;
            val[2] = (typeof val[2] === 'object' ? (val[2] as unknown as CellFormulaValue).result : val[2]) as number;

            const stockSymbol = val[0];
            const buyPrice = val[1];
            const lot = val[2];
            const currentPrice = Number(results[stockSymbol].regularMarketPrice);
            const eps = Number((currentPrice - buyPrice).toPrecision(4));
            const epsp = (eps / buyPrice * 100).toPrecision(4);
            const interestRate = (currentPrice * (results[stockSymbol].dividendYield / 100) / buyPrice * 100).toPrecision(4);
            const totalReturn = eps * lot;
            const totalReturnPercentage = (totalReturn / (buyPrice * lot) * 100).toPrecision(4);

            val.push(currentPrice, eps, `${epsp}%`, `${interestRate}%`, totalReturn, `${totalReturnPercentage}%`);
          });

          wb.removeWorksheet('Stock');
          stockWorksheet = wb.addWorksheet('Stock');

          stockWorksheet.addRow(STOCK_TABLE_HEADERS).commit();

          newTable.stock.forEach((row: Array<CellValue>, rowIndex: number) => {
            const excelRow = stockWorksheet.addRow([
              row[0],
              row[1],
              row[2],
              row[3],
              row[4],
              Number(parseFloat(row[5] as string) / 100),
              Number(parseFloat(row[6] as string) / 100),
              row[7],
              Number(parseFloat(row[8] as string) / 100),
            ]);

            STOCK_TABLE_HEADERS.forEach((_, headerIndex) => {
              const cell = excelRow.getCell(headerIndex + 1);

              if ([1, 3, 4, 7].includes(headerIndex)) {
                cell.numFmt = '$0.00';
              }

              if ([5, 6, 8].includes(headerIndex)) {
                cell.numFmt = '0.00%';
              }

              if (typeof cell.value === 'number' && headerIndex === 8 && cell.value < 0) {
                cell.font = {
                  color: {
                    argb: 'FFFF0000',
                  }
                }
              }
            });

            excelRow.commit();
          });

          newTable.fx.forEach((val: CellValue[]) => {
            val[0] = (typeof val[0] === 'object' ? (val[0] as unknown as CellFormulaValue).result : val[0]) as string;
            val[1] = (typeof val[1] === 'object' ? (val[1] as unknown as CellFormulaValue).result : val[1]) as number;
            val[2] = (typeof val[2] === 'object' ? (val[2] as unknown as CellFormulaValue).result : val[2]) as number;

            const currency = val[0];
            const buyPrice = val[1];
            const lot = val[2];
            const currentPrice = fxRate[currency];
            const total = currentPrice * lot;
            const totalReturn = total - buyPrice * lot;
            const totalReturnPercentage = total === 0 ? 0 : (totalReturn / (buyPrice * lot) * 100).toPrecision(4);

            val.push(currentPrice, total, totalReturn, `${totalReturnPercentage}%`);
          });

          wb.removeWorksheet('Currency');
          fxWorksheet = wb.addWorksheet('Currency');

          fxWorksheet.addRow(FX_TABLE_HEADERS).commit();

          newTable.fx.forEach((row: Array<CellValue>, rowIndex: number) => {
            const excelRow = fxWorksheet.addRow([
              row[0],
              row[1],
              row[2],
              row[3],
              row[4],
              row[5],
              Number(parseFloat(row[6] as string) / 100),
            ]);

            FX_TABLE_HEADERS.forEach((_, headerIndex) => {
              const cell = excelRow.getCell(headerIndex + 1);

              if (typeof cell.value === 'number' && [1, 3, 5].includes(headerIndex)) {
                if (headerIndex === 3) {
                  cell.numFmt = '$0.00000';
                } else {
                  cell.numFmt = '$0.0';
                }
              }

              if (typeof cell.value === 'number' && headerIndex === 6) {
                cell.numFmt = '0.00%';

                if (cell.value < 0) {
                  cell.font = {
                    color: {
                      argb: 'FFFF0000',
                    }
                  }
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

      <div className='sticky top-[100vh]'>V0.7</div>
    </div>
  )
}
