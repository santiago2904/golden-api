import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  app.enableCors(); // Habilitar CORS para que el cliente pueda acceder a los recursos del servidor
  await app.listen(3000);
}
bootstrap();
