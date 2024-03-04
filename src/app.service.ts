import * as XLSX from 'xlsx';
import { Injectable } from '@nestjs/common';
import { Readable } from 'stream';
import axios from 'axios';

export const attendanceData = async () => {
  try {
    const params = {
      user_id: '8722d06a-e334-4450-9793-27ac9578f6a2',
      start_date: '2024-01-15',
      end_date: '2024-02-15',
    };

    const token =
      'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IjNkYjY2YjU1LTUyMDMtNGVjMi05MzQ3LTlkOGNlODcxZTkxNiIsInN0YXR1cyI6MSwiZnVsbF9uYW1lIjoiRmFoYWQgSGltZWwiLCJwaG90byI6IklNR18yMDI0MDEwMl8xMTAxNDQtcmVtb3ZlYmdfXzNfLVBob3RvUm9vbV8xNzA0NTEyNDA3MTAxLnBuZyIsImVtYWlsIjoiZmFoYWRoaW1lbEBnbWFpbC5jb20iLCJwaG9uZV9udW1iZXIiOiIwMTc1MzY0ODI1NiIsImlzX2FjdGl2ZSI6MSwiaXNfYWRtaW4iOjAsInVzZXJfcm9sZXNfaW5mbyI6W3siaWQiOiIwZGY4MGQyZC00MDFmLTRiYTUtYjg0Yy04NTQ3ZDYyM2I0M2MiLCJyb2xlc19pbmZvIjp7ImlkIjoiMmY2ZjFlZTYtNjM4Mi00ZmJjLThjOTQtYjdlYjlhZGQwMWU2IiwibmFtZSI6Imd1ZXN0In19LHsiaWQiOiIzZjU5Y2I0Yy0xZGYxLTQ0MmYtYTE1ZS1iOTZhNGQ4ZTI3MTAiLCJyb2xlc19pbmZvIjp7ImlkIjoiM2NmYTM0M2YtYjYzZS00Nzk2LTljZWYtMTM2N2JiN2FmOGI2IiwibmFtZSI6InN1cGVyX2FkbWluIn19LHsiaWQiOiI3OTA1ZTUyNy0wMTI1LTQ2ZTEtYTllMS00ZTYwYjJkOTVkMjYiLCJyb2xlc19pbmZvIjp7ImlkIjoiYWZkZTQwZTktYzM2ZS00MWVmLTk2NDItMDA5OWZkNTUzZGMwIiwibmFtZSI6InNwZWNpYWxfc3VwZXJfYWRtaW4ifX1dLCJyb2xlX2lkIjoiYWZkZTQwZTktYzM2ZS00MWVmLTk2NDItMDA5OWZkNTUzZGMwIiwicm9sZV9uYW1lIjoic3BlY2lhbF9zdXBlcl9hZG1pbiIsImlhdCI6MTcwNzEyOTU4NSwiZXhwIjoxNzA5NzIxNTg1fQ.XbLHzfJ241DSToGeVT1iOYgxVorWvhLWdRr6YyzII8s';

    const res = await axios.get(
      'http://localhost:8001/api/admin/report/attandence-report/details',
      {
        params,
        headers: {
          Authorization: `Bearer ${token}`,
          // Add other headers if needed
        },
      },
    );
    const { attandence_report_details } = res.data.payload;
    return attandence_report_details;
  } catch (error) {
    console.error(error);
  }
};

@Injectable()
export class AppService {
  getHello(): string {
    return 'Hello World!';
  }

  async createExcelStream(): Promise<Readable> {
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

      // Create a worksheet
      const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet([
        header,
        ...jsonDataObject,
      ]);
      // // Ensure the !ref property is set
      // ws['!ref'] = XLSX.utils.encode_range({
      //   s: { c: 0, r: 0 }, // Start cell (header)
      //   e: { c: header.length - 1, r: jsonDataObject.length }, // End cell
      // });
      // // Style the header
      // const headerRange = XLSX.utils.decode_range(ws['!ref']);
      // console.log('headerRange', headerRange);
      // for (let i = headerRange.s.c; i <= headerRange.e.c; ++i) {
      //   const cellAddress = { r: headerRange.s.r, c: i }; // Header row
      //   const cellRef = XLSX.utils.encode_cell(cellAddress);
      //   const cell = ws[cellRef];

      //   // Add header styling
      //   if (!cell.s) {
      //     cell.s = {};
      //   }

      //   cell.s.font = { bold: true, color: { rgb: '#215325' } }; // Example styles (bold text, red color)
      //   cell.s.alignment = { horizontal: 'center' };
      // }

      // // Center the data in a specific column (let's say column B)
      // const dataRange = XLSX.utils.decode_range(ws['!ref']);
      // for (let i = dataRange.s.r + 1; i <= dataRange.e.r; ++i) {
      //   const cellAddress = { r: i, c: 1 }; // Data in column B (index 1)
      //   const cellRef = XLSX.utils.encode_cell(cellAddress);
      //   const cell = ws[cellRef];

      //   // Add data centering
      //   if (!cell.s) {
      //     cell.s = {};
      //   }

      //   cell.s.alignment = { horizontal: 'center' };
      // }

      const headerWidths = [5, 15, 15, 20, 20, 10, 8, 8, 8]; // Set the desired widths for each column

      // Set column widths
      ws['!cols'] = headerWidths.map((width) => ({ wch: width }));

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

      const stream = new Readable();
      stream.push(excelBuffer);
      stream.push(null); // indicating the end of the stream

      return stream;
    } catch (error) {
      throw error;
    }
  }
}
