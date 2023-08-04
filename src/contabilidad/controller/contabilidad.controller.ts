import { Controller, Post, UseInterceptors, UploadedFiles, Res, Get } from '@nestjs/common';
import { FilesInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { ExcelService } from '../services/excelservice/excel.service';
import { File } from 'multer';


@Controller('contabilidad')
export class ContabilidadController {

    constructor(private readonly excelService: ExcelService) { }

    @Post('facturas')
    @UseInterceptors(FilesInterceptor('archivos', 50))
    async generarExcel(@UploadedFiles() archivos: File[], @Res() res: Response) {
        try {
            const bufferExcel = await this.excelService.generarExcel(archivos);

            await this.excelService.enviarCorreoConAdjunto(
                'Golden.asesorestc@gmail.com', // Reemplaza con el correo de destino
                'Facturas adjuntas', // Asunto del correo
                'Se adjuntan las facturas en formato Excel.', // Cuerpo del correo
                bufferExcel, // Contenido del archivo Excel generado
                'facturas.xlsx', // Nombre del archivo adjunto
            );
            res.set('Content-Disposition', 'attachment; filename=facturas.xlsx');
            res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.send(bufferExcel);
        } catch (error) {
            // Manejo de errores (si es necesario)
            console.error(error);
            res.status(500).send('Error al generar el archivo Excel');
        }
    }

    @Get('get')
    getFacturas() {
        return 'Hola mundo';
    }

}
