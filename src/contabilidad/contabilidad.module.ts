import { Module } from '@nestjs/common';
import { ContabilidadController } from './controller/contabilidad.controller';
import { MulterModule } from '@nestjs/platform-express';
import { ExcelService } from './services/excelservice/excel.service';


@Module({

  imports: [MulterModule.register({
    dest: './uploads'
  })],

  controllers: [ContabilidadController],

  providers: [ExcelService]
})
export class ContabilidadModule { }
