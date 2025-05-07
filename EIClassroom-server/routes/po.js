const express = require('express');
const router = express.Router();
const ExcelJS = require('exceljs');

// Export CO-PO attainment matrix as Excel file
router.post('/export-co-po-matrix', async (req, res) => {
  try {
    const { matrix, averageRow, headers } = req.body;
    
    if (!matrix || !averageRow || !headers) {
      return res.status(400).json({ error: 'Missing required data' });
    }

    // Create a new Excel workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('CO-PO Attainment Matrix');
    
    // Add headers
    const headerRow = [''].concat(headers);
    worksheet.addRow(headerRow);
    
    // Add data rows
    const rowLabels = ["CO1", "CO2", "CO3", "CO4", "CO5", "Average"];
    
    // Add all rows except average
    for (let i = 0; i < matrix.length; i++) {
      const rowData = [rowLabels[i]].concat(matrix[i]);
      worksheet.addRow(rowData);
    }

    // Style the worksheet
    worksheet.getRow(1).font = { bold: true };
    worksheet.getColumn(1).font = { bold: true };
    
    // Get the last row which is the average row
    const lastRow = worksheet.lastRow;
    lastRow.eachCell((cell) => {
      cell.font = { bold: true };
    });

    // Set column widths
    worksheet.columns.forEach(column => {
      column.width = 12;
    });
    worksheet.getColumn(1).width = 15;

    // Set response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=CO_PO_Attainment_Matrix.xlsx');
    
    // Write the workbook to response
    await workbook.xlsx.write(res);
    res.status(200).end();
  } catch (error) {
    console.error('Error generating Excel file:', error);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});

module.exports = router;