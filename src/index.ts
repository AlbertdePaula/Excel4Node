import * as path from 'path';
import Excel from 'exceljs';

const plan = 'olympic-hockey-player.xlsx';

const filePath = path.resolve('./src/data', plan);

type Team = 'M' | 'W';
type Country = 'Canada' | 'USA';
type Position = 'Goalie' | 'Defence' | 'Forward';

type Player = {
    id: number;
    team: Team;
    country: Country;
    firstName: string;
    lastName: string;
    weight: number;
    height: number;
    dateOfBirth: string; // (YYY-MM-DD)
    hometown: string;
    province: string;
    position: Position;
    age: number;
    heightFt: number;
    htln: number;
    bmi: number;
  };

const getCellValue = (row: Excel.Row, cellIndex: number) => {
    const cell = row.getCell(cellIndex);

    return cell.value ? cell.value.toString() : '';
};

const main = async () => {
  const workbook = new Excel.Workbook();
  const content = await workbook.xlsx.readFile(filePath);

  const worksheet = content.worksheets[2]; //Define qual aba da pasta de trabalho;
  const rowStartIndex = 4;  //Define qual linha para iniciar a busca;
  const numberOfRows = worksheet.rowCount - 3;

  const rows = worksheet.getRows(rowStartIndex, numberOfRows) ?? [];

  const players = rows.map((row): Player => {
    return {
      // @ts-ignore
      id: getCellValue(row,1),
      // @ts-ignore
      team: getCellValue(row, 2),
      // @ts-ignore
      country: getCellValue(row, 3),
      firstName: getCellValue(row, 4),
      lastName: getCellValue(row, 5),
      // @ts-ignore
      weight: getCellValue(row, 6),
      height: +getCellValue(row, 7),
      dateOfBirth: getCellValue(row, 8), // (YYY-MM-DD)
      hometown: getCellValue(row, 9),
      province: getCellValue(row, 10),
      // @ts-ignore
      position: getCellValue(row, 11),
      age: +getCellValue(row, 12),
      heightFt: +getCellValue(row, 13),
      htln: +getCellValue(row, 14),
      bmi: +getCellValue(row, 15),
    }
  });

  console.log(players);
};

main().then();

