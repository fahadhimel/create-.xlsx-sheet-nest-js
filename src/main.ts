import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { DocumentBuilder, SwaggerModule } from '@nestjs/swagger';
import { Logger } from '@nestjs/common';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  const PORT = 4005;

  const config = new DocumentBuilder()
    .setTitle('Exal Sheet')
    .setDescription('The Exal Sheet API description')
    .setVersion('1.0')
    // .addTag('Exal Sheet')
    .build();
  const document = SwaggerModule.createDocument(app, config);
  SwaggerModule.setup('api', app, document);

  await app.listen(PORT);
  // Logger
  Logger.log(
    `Server is Running(🔥) on http://127.0.0.1:${PORT}/api/`,
    'Exal sheet',
  );
}
bootstrap();
