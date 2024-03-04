import { Injectable } from '@nestjs/common';
import { Readable } from 'stream';
import { attendanceData } from './app.service';
import * as XLSX from 'xlsx';

@Injectable()
export class ExcelStyleService {
  async style() {
    return 'style';
  }

  async createExcelFileStyle() {
    try {
      const attandence_report_details = await attendanceData();
      // console.log('attandence_report_details', attandence_report_details);

      const jsonDataObject: any[] = [];
      const header = [
        'SL',
        'Date',
        'Day Type',
        'In Time',
        'Out Time',
        'Day',
        'Leaveday',
        'Holiday',
        'Weekend',
      ];

      let index = 1;
      for (const details of attandence_report_details?.user_start_end_date_info) {
        const {
          date,
          day_type,
          start_time,
          end_time,
          day,
          leaveday,
          holiday,
          weekend,
        } = details;

        jsonDataObject.push([
          index,
          date,
          day_type,
          start_time,
          end_time,
          day,
          leaveday,
          holiday,
          weekend,
        ]);
        index++;
      }

      const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet([
        header,
        ...jsonDataObject,
      ]);
      console.log('jsonDataObject', jsonDataObject);
      // Set column widths
      const headerWidths = [5, 15, 15, 20, 20, 10, 8, 8, 8];
      ws['!cols'] = headerWidths.map((width) => ({ wch: width }));

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

      const stream = new Readable();
      stream.push(excelBuffer);
      stream.push(null);

      return stream;
    } catch (error) {
      throw error;
    }
  }

  async generateExcelFile() {
    try {
      const data = [
        [2, '2024-01-16', 'absent', '00:00:00', '00:00:00', 'Tue', 0, 0, 0],
        [3, '2024-01-17', 'holiday', '00:00:00', '00:00:00', 'Wed', 0, 1, 0],
        [4, '2024-01-18', 'absent', '00:00:00', '00:00:00', 'Thu', 0, 0, 0],
      ];

      const headerStyle = {
        font: { bold: true, size: 20, color: { rgb: 'red' } },
        alignment: { horizontal: 'center', vertical: 'center' },
      };

      // Create a new workbook
      const workbook = XLSX.utils.book_new();

      const worksheet = XLSX.utils.aoa_to_sheet(
        [
          [
            { v: 'SL', t: 's', s: headerStyle },
            { v: 'Date', t: 's', s: headerStyle },
            { v: 'Day Type', t: 's', s: headerStyle },
            { v: 'In Time', t: 's', s: headerStyle },
            { v: '', t: 's', s: headerStyle },
            { v: '', t: 's', s: headerStyle },
            { v: 'Out Time', t: 's', s: headerStyle },
            { v: 'Day', t: 's', s: headerStyle },
            { v: 'Leaveday', t: 's', s: headerStyle },
            { v: 'Holiday', t: 's', s: headerStyle },
            { v: 'Weekend', t: 's', s: headerStyle },
          ],
          // Add an empty row for better readability
          [],
          ...data, // Add raw data rows
        ],
        { cellStyles: true },
      );

      // Merge cells for the 'In Time' column header
      worksheet['!merges'] = [{ s: { r: 0, c: 3 }, e: { r: 0, c: 5 } }];

      // Add subheaders for the 'In Time' column
      const inTimeSubheaders = [
        { v: 'Subheader 1', t: 's', s: headerStyle },
        { v: 'Subheader 2', t: 's', s: headerStyle },
        { v: 'Subheader 3', t: 's', s: headerStyle },
      ];
      inTimeSubheaders.forEach((subheader, index) => {
        worksheet[XLSX.utils.encode_cell({ r: 1, c: 3 + index })] = subheader;
      });

      const headerWidths = [5, 15, 15, 20, 20, 10, 8, 8, 8];
      worksheet['!cols'] = headerWidths.map((width) => ({ wch: width }));

      // Add the worksheet to the workbook
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet 1');

      const excelBuffer = XLSX.write(workbook, {
        bookType: 'xlsx',
        type: 'buffer',
        cellStyles: true,
      });

      const stream = new Readable();
      stream.push(excelBuffer);
      stream.push(null);

      return stream;
    } catch (error) {
      throw error;
    }
  }
}
