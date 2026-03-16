const express = require('express');
const cors = require('cors');
const path = require('path'); // Biblioteca obrigatória para o Vercel achar o HTML
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit-table'); 
const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();
const app = express();

app.use(cors());
app.use(express.json());

// Configuração corrigida para o Vercel achar a pasta public e o index.html
app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/api/modulos', async (req, res) => {
    const modulos = await prisma.modulo.findMany({ include: { avaliacoes: true } });
    res.json(modulos);
});

app.post('/api/modulos', async (req, res) => {
    const { nome } = req.body;
    const modulo = await prisma.modulo.create({ data: { nome } });
    res.json(modulo);
});

app.post('/api/avaliacoes', async (req, res) => {
    const data = req.body;
    data.nota = parseFloat(data.nota); 
    const avaliacao = await prisma.avaliacao.create({ data });
    res.json(avaliacao);
});

// --- ROTA DE EXPORTAR EXCEL ---
app.post('/api/export', async (req, res) => {
    const { moduloIds } = req.body; 
    if (!moduloIds || moduloIds.length === 0) return res.status(400).send('Nenhum módulo selecionado.');

    const modulos = await prisma.modulo.findMany({
        where: { id: { in: moduloIds } },
        include: { avaliacoes: true }
    });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Avaliações');

    worksheet.columns = [
        { header: 'TIPO', key: 'tipo', width: 6 },
        { header: 'ASPECTO', key: 'aspecto', width: 45 },
        { header: 'DETALHAMENTO', key: 'detalhe', width: 60 },
        { header: 'NOTA', key: 'nota', width: 10 }
    ];

    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).alignment = { horizontal: 'center' };

    modulos.forEach(modulo => {
        const titleRow = worksheet.addRow([modulo.nome.toUpperCase()]);
        titleRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } };
        worksheet.mergeCells(`A${titleRow.number}:D${titleRow.number}`);

        modulo.avaliacoes.forEach(av => {
            const notaMax = parseFloat(av.nota);
            if (av.tipo === 'M') {
                worksheet.addRow([av.tipo, av.aspecto, av.detalheM, notaMax]);
            } else if (av.tipo === 'J') {
                const nota1 = parseFloat(((notaMax / 3) * 1).toFixed(2));
                const nota2 = parseFloat(((notaMax / 3) * 2).toFixed(2));
                worksheet.addRow([av.tipo, av.aspecto, `0 - ${av.detalhe0}`, 0]);
                worksheet.addRow(['', '', `1 - ${av.detalhe1}`, nota1]);
                worksheet.addRow(['', '', `2 - ${av.detalhe2}`, nota2]);
                worksheet.addRow(['', '', `3 - ${av.detalhe3}`, notaMax]);
            }
        });
        worksheet.addRow([]);
    });

    worksheet.eachRow((row) => {
        row.eachCell({ includeEmpty: true }, (cell) => {
            cell.alignment = { vertical: 'middle', wrapText: true };
            if (row.number > 1 && cell.value !== null && cell.value !== '') {
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            }
        });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Planilha_Avaliacao.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
});

// --- ROTA DE EXPORTAR PDF ---
app.post('/api/export/pdf', async (req, res) => {
    const { moduloIds } = req.body;
    if (!moduloIds || moduloIds.length === 0) return res.status(400).send('Nenhum módulo selecionado.');

    const modulos = await prisma.modulo.findMany({
        where: { id: { in: moduloIds } },
        include: { avaliacoes: true }
    });

    const doc = new PDFDocument({ margin: 30, size: 'A4' });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="Avaliacoes_Senac.pdf"');

    doc.pipe(res);

    doc.fontSize(22).fillColor('#1331a1').text('COMPETIÇÕES SENAC', { align: 'center' });
    doc.fontSize(12).fillColor('#666666').text('Caderno de Avaliação Objetiva e Subjetiva', { align: 'center' });
    doc.moveDown(2);

    for (const modulo of modulos) {
        const tableRows = [];

        modulo.avaliacoes.forEach(av => {
            const notaMax = parseFloat(av.nota);
            if (av.tipo === 'M') {
                tableRows.push([av.tipo, av.aspecto, av.detalheM || '-', notaMax.toFixed(2)]);
            } else if (av.tipo === 'J') {
                const nota1 = ((notaMax / 3) * 1).toFixed(2);
                const nota2 = ((notaMax / 3) * 2).toFixed(2);
                tableRows.push([av.tipo, av.aspecto, `0 - ${av.detalhe0 || '-'}`, "0.00"]);
                tableRows.push(['', '', `1 - ${av.detalhe1 || '-'}`, nota1]);
                tableRows.push(['', '', `2 - ${av.detalhe2 || '-'}`, nota2]);
                tableRows.push(['', '', `3 - ${av.detalhe3 || '-'}`, notaMax.toFixed(2)]);
            }
        });

        const tableConfig = {
            title: `Módulo: ${modulo.nome.toUpperCase()}`,
            headers: [
                { label: "TIPO", property: 'tipo', width: 40 },
                { label: "ASPECTO / CRITÉRIO", property: 'aspecto', width: 160 },
                { label: "DETALHAMENTO (RUBRICA)", property: 'detalhe', width: 280 },
                { label: "NOTA", property: 'nota', width: 50 }
            ],
            rows: tableRows
        };

        await doc.table(tableConfig, {
            prepareHeader: () => doc.font("Helvetica-Bold").fontSize(9).fillColor('#1331a1'),
            prepareRow: (row, indexColumn, indexRow, rectRow, rectCell) => {
                doc.font("Helvetica").fontSize(9).fillColor('#333333');
            }
        });
        
        doc.moveDown(1);
    }

    doc.end();
});

app.put('/api/avaliacoes/:id', async (req, res) => {
    const { id } = req.params;
    const data = req.body;
    if (data.nota) data.nota = parseFloat(data.nota);
    try {
        const avaliacao = await prisma.avaliacao.update({ where: { id: parseInt(id) }, data });
        res.json(avaliacao);
    } catch (error) { res.status(500).json({ error: "Erro ao atualizar." }); }
});

app.delete('/api/avaliacoes/:id', async (req, res) => {
    const { id } = req.params;
    try {
        await prisma.avaliacao.delete({ where: { id: parseInt(id) } });
        res.json({ message: "Excluída com sucesso." });
    } catch (error) { res.status(500).json({ error: "Erro ao excluir." }); }
});

app.delete('/api/modulos/:id', async (req, res) => {
    const { id } = req.params;
    try {
        await prisma.avaliacao.deleteMany({ where: { moduloId: parseInt(id) } });
        await prisma.modulo.delete({ where: { id: parseInt(id) } });
        res.json({ message: "Módulo excluído com sucesso." });
    } catch (error) { res.status(500).json({ error: "Erro ao excluir o módulo." }); }
});

// Exporta o app para o Vercel conseguir rodar
module.exports = app;

// Mantém o app.listen apenas se você for rodar localmente no seu computador
if (require.main === module) {
    app.listen(3000, () => console.log('Servidor rodando na porta 3000'));
}