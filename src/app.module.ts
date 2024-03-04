import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ExcelStyleService } from './excel-style.service';

@Module({
  imports: [],
  controllers: [AppController],
  providers: [AppService, ExcelStyleService],
})
export class AppModule {}
