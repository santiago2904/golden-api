
import { Injectable, HttpException, HttpStatus } from '@nestjs/common';
import { Workbook, Worksheet } from 'exceljs';
import * as fs from 'fs';
import { File } from 'multer';
import * as nodemailer from 'nodemailer';

@Injectable()
export class ExcelService {
    async generarExcel(archivos: File[]) {
        if (!archivos || archivos.length === 0) {
            throw new HttpException('No se recibieron archivos XML', HttpStatus.BAD_REQUEST);
        }

        console.log(archivos.length);
        const libroExcel = new Workbook();
        const hojaExcel = libroExcel.addWorksheet('FacturasVenta');

        // Definir los encabezados de las columnas en el archivo Excel
        const encabezados = [
            'Archivo',
            'Número',
            'Empresa',
            'NIT',
            'LineExtensionAmount',
            'TaxExclusiveAmount',
            'TaxInclusiveAmount',
            'IVA',
            'Fecha de emisión',
            'Fecha de vencimiento',
        ];
        hojaExcel.addRow(encabezados);

        // Procesar los archivos XML y agregar los datos al archivo Excel
        for (let i = 0; i < archivos.length; i++) {
            const archivo = archivos[i];
            const contenidoXml = fs.readFileSync(archivo.path, 'utf-8'); // Leer el contenido XML del archivo

            // Buscar y extraer los datos del contenido XML
            const nombreEmpresa = this.obtenerValorEtiquetaXml(contenidoXml, '<cbc:Name>(.*?)<\/cbc:Name>');
            const nitEmpresa = this.obtenerValorEtiquetaXml(contenidoXml, '<cbc:CompanyID.*?>(.*?)<\/cbc:CompanyID>');
            const lineExtensionAmount = this.obtenerValorEtiquetaXml(contenidoXml, '<cbc:LineExtensionAmount.*?>(.*?)<\/cbc:LineExtensionAmount>');
            const taxExclusiveAmount = this.obtenerValorEtiquetaXml(contenidoXml, '<cbc:TaxExclusiveAmount.*?>(.*?)<\/cbc:TaxExclusiveAmount>');
            const taxInclusiveAmount = this.obtenerValorEtiquetaXml(contenidoXml, '<cbc:TaxInclusiveAmount.*?>(.*?)<\/cbc:TaxInclusiveAmount>');
            const iva = this.obtenerValorEtiquetaXml(contenidoXml, '<cbc:TaxAmount.*?>(.*?)<\/cbc:TaxAmount>');
            const fechaEmision = this.obtenerValorEtiquetaXml(contenidoXml, '<cbc:IssueDate>(.*?)<\/cbc:IssueDate>');
            const fechaVencimiento = this.obtenerValorEtiquetaXml(contenidoXml, '<cbc:DueDate>(.*?)<\/cbc:DueDate>');

            // Agregar los datos al archivo Excel
            hojaExcel.addRow([
                archivo.originalname,
                i + 1,
                nombreEmpresa,
                nitEmpresa,
                lineExtensionAmount,
                taxExclusiveAmount,
                taxInclusiveAmount,
                iva,
                fechaEmision,
                fechaVencimiento,
            ]);
        }

        // Eliminar los archivos temporales de los XML después de procesarlos
        this.eliminarArchivosTemporales(archivos);

        // Generar el archivo de Excel y devolver su contenido en forma de Buffer
        const bufferExcel = await libroExcel.xlsx.writeBuffer();
        return Buffer.from(bufferExcel) as Buffer;
    }

    // Método para eliminar los archivos temporales de los XML después de procesarlos
    private eliminarArchivosTemporales(archivos: File[]): void {
        archivos.forEach((archivo) => {
            fs.unlinkSync(archivo.path);
        });
    }

    // Método para obtener el valor de una etiqueta en un contenido XML utilizando expresiones regulares
    private obtenerValorEtiquetaXml(xml: string, etiqueta: string): string {
        const regex = new RegExp(etiqueta, 's'); // 's' indica que el punto (.) también debe coincidir con saltos de línea
        const match = xml.match(regex);
        return match ? match[1].trim() : '';
    }


    async enviarCorreoConAdjunto(correoDestino: string, asunto: string, cuerpo: string, archivoAdjunto: Buffer, nombreArchivo: string) {
        try {
            // Configurar el transporte de nodemailer (utilizando una cuenta de Gmail)
            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: {
                    user: 'santiagopalacioalzate@gmail.com', // Reemplaza con tu correo de Gmail
                    pass: 'edubbqmlkgicayyv', // Reemplaza con tu contraseña de Gmail
                },
            });

            // Configurar el correo electrónico
            const mailOptions = {
                from: 'santiagopalacioalzate@gmail.com', // Reemplaza con tu correo de Gmail
                to: correoDestino,
                subject: asunto,
                text: cuerpo,
                attachments: [
                    {
                        filename: nombreArchivo,
                        content: archivoAdjunto,
                    },
                ],
            };

            // Enviar el correo
            const info = await transporter.sendMail(mailOptions);
            console.log('Correo enviado: ', info.response);
        } catch (error) {
            console.error('Error al enviar el correo:', error);
            throw new Error('Error al enviar el correo');
        }
    }


}
