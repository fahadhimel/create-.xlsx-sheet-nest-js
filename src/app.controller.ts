import { Controller, Get, Res } from '@nestjs/common';
import { AppService } from './app.service';
import { ApiTags } from '@nestjs/swagger';
import { Response } from 'express';

@ApiTags('exal-sheet')
@Controller('exal-sheet')
export class AppController {
  constructor(private readonly appService: AppService) {}

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
}
