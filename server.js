const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();
const app = express();

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// 1. Rota para listar módulos e avaliações
app.get('/api/modulos', async (req, res) => {
    const modulos = await prisma.modulo.findMany({ include: { avaliacoes: true } });
    res.json(modulos);
});

// 2. Rota para criar um Módulo
app.post('/api/modulos', async (req, res) => {
    const { nome } = req.body;
    const modulo = await prisma.modulo.create({ data: { nome } });
    res.json(modulo);
});

// 3. Rota para criar uma Avaliação
app.post('/api/avaliacoes', async (req, res) => {
    const data = req.body;
    data.nota = parseFloat(data.nota); 
    const avaliacao = await prisma.avaliacao.create({ data });
    res.json(avaliacao);
});

// 4. Rota de Exportação Excel
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
                cell.border = {
                    top: { style: 'thin' }, left: { style: 'thin' },
                    bottom: { style: 'thin' }, right: { style: 'thin' }
                };
            }
        });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Planilha_Avaliacao.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
});

// 5. Rota para Atualizar Avaliação
app.put('/api/avaliacoes/:id', async (req, res) => {
    const { id } = req.params;
    const data = req.body;
    if (data.nota) data.nota = parseFloat(data.nota);
    try {
        const avaliacao = await prisma.avaliacao.update({ where: { id: parseInt(id) }, data });
        res.json(avaliacao);
    } catch (error) { res.status(500).json({ error: "Erro ao atualizar." }); }
});

// 6. Rota para Deletar Avaliação
app.delete('/api/avaliacoes/:id', async (req, res) => {
    const { id } = req.params;
    try {
        await prisma.avaliacao.delete({ where: { id: parseInt(id) } });
        res.json({ message: "Excluída com sucesso." });
    } catch (error) { res.status(500).json({ error: "Erro ao excluir." }); }
});

// 7. Rota para Deletar Módulo Inteiro (e todas as suas avaliações)
app.delete('/api/modulos/:id', async (req, res) => {
    const { id } = req.params;
    try {
        // Primeiro deleta todas as avaliações que pertencem a este módulo
        await prisma.avaliacao.deleteMany({ where: { moduloId: parseInt(id) } });
        // Depois deleta o módulo em si
        await prisma.modulo.delete({ where: { id: parseInt(id) } });
        res.json({ message: "Módulo excluído com sucesso." });
    } catch (error) {
        res.status(500).json({ error: "Erro ao excluir o módulo." });
    }
});

app.listen(3000, () => console.log('Servidor rodando na porta 3000'));