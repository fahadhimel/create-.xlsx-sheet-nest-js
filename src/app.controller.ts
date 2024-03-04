import { Controller, Get, Res } from '@nestjs/common';
import { AppService } from './app.service';
import { ApiTags } from '@nestjs/swagger';
import { Response } from 'express';
import { ExcelStyleService } from './excel-style.service';

@ApiTags('exal-sheet')
@Controller('exal-sheet')
export class AppController {
  constructor(
    private readonly appService: AppService,
    private readonly excelStyleService: ExcelStyleService,
  ) {}

  @Get()
  getHello(): string {
    return this.appService.getHello();
  }

  @Get('create')
  async createExcelFile(@Res() res: Response) {
    const stream = await this.appService.createExcelStream();

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="file.xlsx"',
    });

    stream.pipe(res);
  }

  @Get('style')
  async Style(): Promise<string> {
    return this.excelStyleService.style();
  }

  @Get('style/create')
  async createExcelFileStyle(@Res() res: Response) {
    const stream = await this.excelStyleService.createExcelFileStyle();

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="file.xlsx"',
    });

    stream.pipe(res);
  }

  @Get('generate-excel')
  async generateExcel(@Res() res: Response) {
    const stream = await this.excelStyleService.generateExcelFile();

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="file.xlsx"',
    });

    stream.pipe(res);
  }
}
