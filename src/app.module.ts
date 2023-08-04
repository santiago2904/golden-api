import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';

import { ContabilidadModule } from './contabilidad/contabilidad.module';

@Module({
  imports: [ContabilidadModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
