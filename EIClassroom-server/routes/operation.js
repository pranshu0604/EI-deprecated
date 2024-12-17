const express = require('express');
const { PrismaClient } = require('@prisma/client');
const router = express.Router();

router.use(express.json());

const prisma = new PrismaClient();

// Route to post Exam CO schema
router.post('/co-form', async (req, res) => {
  const { subjectCode, mst1, mst2 } = req.body;

  try {
    const newCO = await prisma.cO.upsert({
      where: { subjectCode },
      update: {
        MST1_Q1: mst1.Q1,
        MST1_Q2: mst1.Q2,
        MST1_Q3: mst1.Q3,
        MST2_Q1: mst2.Q1,
        MST2_Q2: mst2.Q2,
        MST2_Q3: mst2.Q3,
      },
      create: {
        subjectCode,
        MST1_Q1: mst1.Q1,
        MST1_Q2: mst1.Q2,
        MST1_Q3: mst1.Q3,
        MST2_Q1: mst2.Q1,
        MST2_Q2: mst2.Q2,
        MST2_Q3: mst2.Q3,
      },
    });

    res.status(201).json(newCO);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Failed to submit the form' });
  }
});

// Route to create a new sheet entry
router.post('/submit-form', async (req, res) => {
  const { id, name, subjectCode, mst1, mst2, assignment, endsem } = req.body;

  const MST1_Q1 = mst1?.Q1;
  const MST1_Q2 = mst1?.Q2;
  const MST1_Q3 = mst1?.Q3;

  const MST2_Q1 = mst2?.Q1;
  const MST2_Q2 = mst2?.Q2;
  const MST2_Q3 = mst2?.Q3;

  const Assignment_CO1 = assignment?.CO1;
  const Assignment_CO2 = assignment?.CO2;
  const Assignment_CO3 = assignment?.CO3;
  const Assignment_CO4 = assignment?.CO4;
  const Assignment_CO5 = assignment?.CO5;

  const EndSem_Q1 = endsem?.Q1;
  const EndSem_Q2 = endsem?.Q2;
  const EndSem_Q3 = endsem?.Q3;
  const EndSem_Q4 = endsem?.Q4;
  const EndSem_Q5 = endsem?.Q5;

  try {
    // Find the teacher based on subjectCode
    const subject = await prisma.subject.findUnique({
      where: {
        code: subjectCode,
      },
      include: {
        teacher: true, // Include the teacher data in the result
      },
    });

    if (!subject || !subject.teacher) {
      return res.status(404).json({ error: 'Subject or Teacher not found' });
    }

    // Create the new sheet and connect it with the subject and teacher
    await prisma.sheet.create({
      data: {
        id,
        name,
        subjectCode,
        teacherId: subject.teacher.id, // Dynamically connect teacher through subjectCode
        MST1_Q1: parseFloat(MST1_Q1),
        MST1_Q2: parseFloat(MST1_Q2),
        MST1_Q3: parseFloat(MST1_Q3),
        MST2_Q1: parseFloat(MST2_Q1),
        MST2_Q2: parseFloat(MST2_Q2),
        MST2_Q3: parseFloat(MST2_Q3),
        EndSem_Q1: parseFloat(EndSem_Q1),
        EndSem_Q2: parseFloat(EndSem_Q2),
        EndSem_Q3: parseFloat(EndSem_Q3),
        EndSem_Q4: parseFloat(EndSem_Q4),
        EndSem_Q5: parseFloat(EndSem_Q5),
        Assignment_CO1: parseFloat(Assignment_CO1),
        Assignment_CO2: parseFloat(Assignment_CO2),
        Assignment_CO3: parseFloat(Assignment_CO3),
        Assignment_CO4: parseFloat(Assignment_CO4),
        Assignment_CO5: parseFloat(Assignment_CO5),
      },
    });

    res.status(201).json({ message: 'Form data saved successfully' });
  } catch (error) {
    console.error('Error saving form data:', error);
    res.status(500).json({ error: 'Error saving form data' });
  }
});

// Route to delete a student's record by ID and subjectCode
router.delete('/sheets/:id/:subjectCode', async (req, res) => {
  const { id, subjectCode } = req.params;

  try {
    await prisma.sheet.delete({
      where: {
        id_subjectCode: {
          id: id,
          subjectCode: subjectCode,
        },
      },
    });

    res.status(200).json({ message: 'Student record deleted successfully' });
  } catch (error) {
    console.error('Error deleting student record:', error);
    res.status(500).json({ error: 'Error deleting student record' });
  }
});

// Route to get all rows from the Sheet table with a specific subjectCode
router.get('/sheets', async (req, res) => {
  const { subjectCode } = req.query; // Retrieve subjectCode from query parameter

  try {
    const sheets = await prisma.sheet.findMany({
      where: {
        subjectCode: subjectCode, // Filter records based on subjectCode
      },
    });
    res.status(200).json(sheets);
  } catch (error) {
    console.error('Error fetching sheets:', error);
    res.status(500).json({ error: 'Error fetching sheets' });
  }
});

// Route to update a specific sheet entry by ID and subjectCode
router.put('/sheets/:id/:subjectCode', async (req, res) => {
  const { id, subjectCode } = req.params;
  const {         
    name,
    MST1_Q1,
    MST1_Q2,
    MST1_Q3,
    MST2_Q1,
    MST2_Q2,
    MST2_Q3,
    Assignment_CO1,
    Assignment_CO2,
    Assignment_CO3,
    Assignment_CO4,
    Assignment_CO5,
    EndSem_Q1,
    EndSem_Q2,
    EndSem_Q3,
    EndSem_Q4,
    EndSem_Q5 
  } = req.body;

  try {
    const updatedSheet = await prisma.sheet.update({
      where: {
        id_subjectCode: {
          id: id,
          subjectCode: subjectCode,
        },
      },
      data: {
        name,
        MST1_Q1,
        MST1_Q2,
        MST1_Q3,
        MST2_Q1,
        MST2_Q2,
        MST2_Q3,
        Assignment_CO1,
        Assignment_CO2,
        Assignment_CO3,
        Assignment_CO4,
        Assignment_CO5,
        EndSem_Q1,
        EndSem_Q2,
        EndSem_Q3,
        EndSem_Q4,
        EndSem_Q5
      },
    });

    res.status(200).json(updatedSheet);
  } catch (error) {
    console.error('Error updating sheet:', error);
    res.status(500).json({ error: 'Error updating sheet' });
  }
});

const ExcelJS = require('exceljs');

router.get('/downloadmst1/:subjectCode', async (req, res) => {
  const { subjectCode } = req.params;

  try {
    // Fetch CO mappings from the CO table
    const coData = await prisma.cO.findUnique({
      where: { subjectCode },
    });

    if (!coData) {
      return res.status(404).json({ error: 'CO mapping not found for this subject' });
    }

    // Fetch student scores from the Sheet table
    const studentScores = await prisma.sheet.findMany({
      where: { subjectCode },
    });

    if (studentScores.length === 0) {
      return res.status(404).json({ error: 'No student scores found for this subject' });
    }

    const studentCount = studentScores.length;

    // Create a new Excel workbook and sheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('CO Attainment');

    // Get unique mapped COs for MST1
    const mappedCOs = [...new Set([coData.MST1_Q1, coData.MST1_Q2, coData.MST1_Q3])].filter(Boolean);

    // Set column widths and headers
    const columns = [
      { header: 'Enrollment Number', key: 'enrollment', width: 20 },
      { header: 'Name', key: 'name', width: 30 },
      { header: `Q1-${coData.MST1_Q1}`, key: 'q1', width: 15 },
      { header: `Q2-${coData.MST1_Q2}`, key: 'q2', width: 15 },
      { header: `Q3-${coData.MST1_Q3}`, key: 'q3', width: 15 }
    ];

    // Add Total CO columns only for mapped COs
    mappedCOs.forEach(co => {
      columns.push({ header: `Total ${co}`, key: co, width: 15 });
    });

    worksheet.columns = columns;

    // Function to calculate CO totals for a student
    const calculateCOTotals = (student) => {
      const totals = {};
      mappedCOs.forEach(co => totals[co] = null);

      if (coData.MST1_Q1 && student.MST1_Q1 != null) totals[coData.MST1_Q1] = (totals[coData.MST1_Q1] || 0) + parseFloat(student.MST1_Q1);
      if (coData.MST1_Q2 && student.MST1_Q2 != null) totals[coData.MST1_Q2] = (totals[coData.MST1_Q2] || 0) + parseFloat(student.MST1_Q2);
      if (coData.MST1_Q3 && student.MST1_Q3 != null) totals[coData.MST1_Q3] = (totals[coData.MST1_Q3] || 0) + parseFloat(student.MST1_Q3);

      return totals;
    };

    // Initialize grand totals and counts
    const grandTotals = {};
    const counts = {};
    mappedCOs.forEach(co => {
      grandTotals[co] = 0;
      counts[co] = 0;
    });

    // Add rows for each student with their scores
    studentScores.forEach((student) => {
      const coTotals = calculateCOTotals(student);
      
      // Add to grand totals and counts
      mappedCOs.forEach(co => {
        if (coTotals[co] != null) {
          grandTotals[co] += coTotals[co];
          counts[co]++;
        }
      });

      worksheet.addRow({
        enrollment: student.id,
        name: student.name,
        q1: student.MST1_Q1 != null ? parseFloat(student.MST1_Q1) : '',
        q2: student.MST1_Q2 != null ? parseFloat(student.MST1_Q2) : '',
        q3: student.MST1_Q3 != null ? parseFloat(student.MST1_Q3) : '',
        ...coTotals
      });
    });

    // Calculate and add averages row
    const averages = {};
    mappedCOs.forEach(co => averages[co] = counts[co] ? Math.floor(grandTotals[co] / counts[co]) : null);

    const targetRow = worksheet.addRow({
      enrollment: 'Average (Target Marks)',
      name: '',
      q1: '',
      q2: '',
      q3: '',
      ...averages
    });

    targetRow.font = { bold: true };
    targetRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFD700' }
    };

    // Count students above target
    const studentsAboveTarget = {};
    mappedCOs.forEach(co => studentsAboveTarget[co] = 0);

    studentScores.forEach(student => {
      const coTotals = calculateCOTotals(student);
      mappedCOs.forEach(co => {
        if (coTotals[co] != null && coTotals[co] >= averages[co]) {
          studentsAboveTarget[co]++;
        }
      });
    });

    // Calculate and add percentages row
    const percentages = {};
    mappedCOs.forEach(co => {
      // Only count students who have a non-null value for this CO
      const validStudentCount = studentScores.filter(student => {
        const coTotal = calculateCOTotals(student)[co];
        return coTotal != null;
      }).length;

      if (validStudentCount > 0) {
        percentages[co] = `${((studentsAboveTarget[co] / validStudentCount) * 100).toFixed(2)}%`;
      } else {
        percentages[co] = '';
      }
    });

    worksheet.addRow({
      enrollment: 'Percentage',
      name: '',
      q1: '',
      q2: '',
      q3: '',
      ...percentages
    });

    // Calculate and add CO levels row
    const coLevels = {};
    mappedCOs.forEach(co => {
      const percentage = parseFloat(percentages[co]);
      coLevels[co] = percentage >= 70 ? 3 : percentage >= 60 ? 2 : percentage >= 50 ? 1 : 0;
    });

    const coLevelRow = worksheet.addRow({
      enrollment: 'CO Level',
      name: '',
      q1: '',
      q2: '',
      q3: '',
      ...coLevels
    });

    coLevelRow.font = { bold: true };
    coLevelRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF98FB98' }
    };

    // Prepare the response with the generated Excel file
    const fileName = `CO_Attainment_${subjectCode}.xlsx`;
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Write the workbook to the response
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Failed to generate the Excel sheet' });
  }
});

router.get('/downloadmst2/:subjectCode', async (req, res) => {
  const { subjectCode } = req.params;

  try {
    // Fetch CO mappings from the CO table
    const coData = await prisma.cO.findUnique({
      where: { subjectCode },
    });

    if (!coData) {
      return res.status(404).json({ error: 'CO mapping not found for this subject' });
    }

    // Fetch student scores from the Sheet table
    const studentScores = await prisma.sheet.findMany({
      where: { subjectCode },
    });

    if (studentScores.length === 0) {
      return res.status(404).json({ error: 'No student scores found for this subject' });
    }

    const studentCount = studentScores.length;

    // Create a new Excel workbook and sheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('CO Attainment');

    // Get unique mapped COs for MST2
    const mappedCOs = [...new Set([coData.MST2_Q1, coData.MST2_Q2, coData.MST2_Q3])].filter(Boolean);

    // Set column widths and headers
    const columns = [
      { header: 'Enrollment Number', key: 'enrollment', width: 20 },
      { header: 'Name', key: 'name', width: 30 },
      { header: `Q1-${coData.MST2_Q1}`, key: 'q1', width: 15 },
      { header: `Q2-${coData.MST2_Q2}`, key: 'q2', width: 15 },
      { header: `Q3-${coData.MST2_Q3}`, key: 'q3', width: 15 }
    ];

    // Add Total CO columns only for mapped COs
    mappedCOs.forEach(co => {
      columns.push({ header: `Total ${co}`, key: co, width: 15 });
    });

    worksheet.columns = columns;

    // Function to calculate CO totals for a student
    const calculateCOTotals = (student) => {
      const totals = {};
      mappedCOs.forEach(co => totals[co] = null);

      if (coData.MST2_Q1 && student.MST2_Q1 != null) totals[coData.MST2_Q1] = (totals[coData.MST2_Q1] || 0) + parseFloat(student.MST2_Q1);
      if (coData.MST2_Q2 && student.MST2_Q2 != null) totals[coData.MST2_Q2] = (totals[coData.MST2_Q2] || 0) + parseFloat(student.MST2_Q2);
      if (coData.MST2_Q3 && student.MST2_Q3 != null) totals[coData.MST2_Q3] = (totals[coData.MST2_Q3] || 0) + parseFloat(student.MST2_Q3);

      return totals;
    };

    // Initialize grand totals and counts
    const grandTotals = {};
    const counts = {};
    mappedCOs.forEach(co => {
      grandTotals[co] = 0;
      counts[co] = 0;
    });

    // Add rows for each student with their scores
    studentScores.forEach((student) => {
      const coTotals = calculateCOTotals(student);
      
      // Add to grand totals and counts
      mappedCOs.forEach(co => {
        if (coTotals[co] != null) {
          grandTotals[co] += coTotals[co];
          counts[co]++;
        }
      });

      worksheet.addRow({
        enrollment: student.id,
        name: student.name,
        q1: student.MST2_Q1 != null ? parseFloat(student.MST2_Q1) : '',
        q2: student.MST2_Q2 != null ? parseFloat(student.MST2_Q2) : '',
        q3: student.MST2_Q3 != null ? parseFloat(student.MST2_Q3) : '',
        ...coTotals
      });
    });

    // Calculate and add averages row
    const averages = {};
    mappedCOs.forEach(co => averages[co] = counts[co] ? Math.floor(grandTotals[co] / counts[co]) : null);

    const targetRow = worksheet.addRow({
      enrollment: 'Average (Target Marks)',
      name: '',
      q1: '',
      q2: '',
      q3: '',
      ...averages
    });

    targetRow.font = { bold: true };
    targetRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFD700' }
    };

    // Count students above target
    const studentsAboveTarget = {};
    mappedCOs.forEach(co => studentsAboveTarget[co] = 0);

    studentScores.forEach(student => {
      const coTotals = calculateCOTotals(student);
      mappedCOs.forEach(co => {
        if (coTotals[co] != null && coTotals[co] >= averages[co]) {
          studentsAboveTarget[co]++;
        }
      });
    });

    // Calculate and add percentages row
    const percentages = {};
    mappedCOs.forEach(co => {
      // Only count students who have a non-null value for this CO
      const validStudentCount = studentScores.filter(student => {
        const coTotal = calculateCOTotals(student)[co];
        return coTotal != null;
      }).length;

      if (validStudentCount > 0) {
        percentages[co] = `${((studentsAboveTarget[co] / validStudentCount) * 100).toFixed(2)}%`;
      } else {
        percentages[co] = '';
      }
    });

    worksheet.addRow({
      enrollment: 'Percentage',
      name: '',
      q1: '',
      q2: '',
      q3: '',
      ...percentages
    });

    // Calculate and add CO levels row
    const coLevels = {};
    mappedCOs.forEach(co => {
      const percentage = parseFloat(percentages[co]);
      coLevels[co] = percentage >= 70 ? 3 : percentage >= 60 ? 2 : percentage >= 50 ? 1 : 0;
    });

    const coLevelRow = worksheet.addRow({
      enrollment: 'CO Level',
      name: '',
      q1: '',
      q2: '',
      q3: '',
      ...coLevels
    });

    coLevelRow.font = { bold: true };
    coLevelRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF98FB98' }
    };

    // Prepare the response with the generated Excel file
    const fileName = `CO_Attainment_${subjectCode}.xlsx`;
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Write the workbook to the response
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Failed to generate the Excel sheet' });
  }
});

// Route to download an Excel sheet with calculations and averages
router.get('/overall-sheet', async (req, res) => {
  const { subjectCode } = req.query; // Get subjectCode from query parameter

  try {
    // Fetch the subject details
    const subject = await prisma.subject.findUnique({
      where: { code: subjectCode },
    });

    if (!subject) {
      return res.status(404).json({ error: 'Subject not found.' });
    }

    // Fetch the sheets for the specified subjectCode
    const sheets = await prisma.sheet.findMany({
      where: {
        subjectCode: subjectCode,
      },
    });

    if (sheets.length === 0) {
      return res.status(404).json({ error: 'No sheets found for this subject code.' });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet Data');

    // Add header rows
    worksheet.addRow(['Shri G. S. Institute of Tech. and Science']);
    worksheet.addRow(['Department of Electronics and Instrumentation Engineering']);
    worksheet.addRow([`Course Outcome Sheet for ${subjectCode}-${subject.name}`]);

    // Center align and merge the header rows
    for (let i = 1; i <= 3; i++) {
      worksheet.getRow(i).alignment = { horizontal: 'center' };
      worksheet.mergeCells(`A${i}:S${i}`);
    }

    // Add an empty row for spacing
    worksheet.addRow([]);

    // Add main header row
    worksheet.addRow([
      'Enrollment Number', 'Name', 'Subject Code', 'MST1_Q1', 'MST1_Q2', 'MST1_Q3', 'MST1_Total',
      'MST2_Q1', 'MST2_Q2', 'MST2_Q3', 'MST2_Total', 'MST_Best', 'Assignment_CO1', 'Assignment_CO2', 'Assignment_CO3', 'Assignment_CO4', 'Assignment_CO5', 'Assignment_Total',
      'EndSem_Q1', 'EndSem_Q2', 'EndSem_Q3', 'EndSem_Q4', 'EndSem_Q5', 'EndSem_Total'
    ]);

    // Style the main header row
    worksheet.getRow(5).font = { bold: true };
    worksheet.getRow(5).alignment = { horizontal: 'center' };

    // Set column widths
    worksheet.columns = [
      { width: 20 }, { width: 30 }, { width: 15 },
      { width: 10 }, { width: 10 }, { width: 10 }, { width: 15 },
      { width: 10 }, { width: 10 }, { width: 10 }, { width: 15 },
      { width: 15 },
      { width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }, { width: 20 },
      { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 15 }
    ];

    // Add data rows with calculations for MST1_Total, MST2_Total, MST_Best, and EndSem_Total
    sheets.forEach((sheet) => {
      const MST1_Total = (sheet.MST1_Q1 || 0) + (sheet.MST1_Q2 || 0) + (sheet.MST1_Q3 || 0);
      const MST2_Total = (sheet.MST2_Q1 || 0) + (sheet.MST2_Q2 || 0) + (sheet.MST2_Q3 || 0);
      const MST_Best = Math.max(MST1_Total, MST2_Total);
      const Assignment_Total = (sheet.Assignment_CO1 || 0) + (sheet.Assignment_CO2 || 0) + (sheet.Assignment_CO3 || 0) + (sheet.Assignment_CO4 || 0) + (sheet.Assignment_CO5 || 0);
      const EndSem_Total = (sheet.EndSem_Q1 || 0) + (sheet.EndSem_Q2 || 0) + (sheet.EndSem_Q3 || 0) + (sheet.EndSem_Q4 || 0) + (sheet.EndSem_Q5 || 0);

      worksheet.addRow([
        sheet.id, sheet.name, sheet.subjectCode,
        sheet.MST1_Q1, sheet.MST1_Q2, sheet.MST1_Q3, MST1_Total,
        sheet.MST2_Q1, sheet.MST2_Q2, sheet.MST2_Q3, MST2_Total,
        MST_Best,
        sheet.Assignment_CO1, sheet.Assignment_CO2, sheet.Assignment_CO3, sheet.Assignment_CO4, sheet.Assignment_CO5, Assignment_Total,
        sheet.EndSem_Q1, sheet.EndSem_Q2, sheet.EndSem_Q3, sheet.EndSem_Q4, sheet.EndSem_Q5, EndSem_Total
      ]);
    });

    // Add the average row at the end of the data
    const lastRowNumber = worksheet.lastRow.number;
    const averageRow = worksheet.addRow(['Average']);
    averageRow.font = { bold: true };

    // Calculate averages for each column
    for (let col = 4; col <= 24; col++) {
      averageRow.getCell(col).value = { formula: `AVERAGE(${String.fromCharCode(64 + col)}6:${String.fromCharCode(64 + col)}${lastRowNumber})` };
    }

    // Set response headers for the download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=sheets.xlsx');

    // Write the Excel file to the response
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error generating Excel sheet:', error);
    res.status(500).json({ error: 'Error generating Excel sheet' });
  }
});

// Add this route after the other routes
router.get('/co-matrix/:subjectCode', async (req, res) => {
  try {
    const { subjectCode } = req.params;
    
    // Fetch CO mappings
    const coMapping = await prisma.cO.findUnique({
      where: { subjectCode }
    });

    if (!coMapping) {
      return res.status(404).json({ error: 'CO mapping not found' });
    }

    // Fetch all sheets for calculations
    const sheets = await prisma.sheet.findMany({
      where: { subjectCode }
    });

    if (sheets.length === 0) {
      return res.status(404).json({ error: 'No sheets found' });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('CO Matrix');

    // Helper function to calculate CO level based on target score
    const calculateCOLevel = (scores) => {
      if (!scores.length) return '';
      
      // Step 1: Calculate average from valid scores only
      const validScores = scores
        .filter(score => score !== null && score !== undefined)
        .map(score => parseFloat(score));

      if (validScores.length === 0) return '';
      
      const average = Math.floor(validScores.reduce((sum, score) => sum + score, 0) / validScores.length);
      
      // Step 2: This average is our target score
      // Step 3: Calculate percentage based on ALL students
      const studentsAboveTarget = scores
        .filter(score => score !== null && score !== undefined && parseFloat(score) >= average)
        .length;
      const percentage = (studentsAboveTarget / scores.length) * 100; // Using total students count

      if (percentage >= 70) return 3;
      if (percentage >= 60) return 2;
      if (percentage >= 50) return 1;
      return 0;
    };

    // Set up headers
    worksheet.columns = [
      { header: 'Assessment', key: 'assessment', width: 15 },
      { header: 'CO1', key: 'co1', width: 10 },
      { header: 'CO2', key: 'co2', width: 10 },
      { header: 'CO3', key: 'co3', width: 10 },
      { header: 'CO4', key: 'co4', width: 10 },
      { header: 'CO5', key: 'co5', width: 10 }
    ];

    // Style headers
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' }
    };

    // Calculate MST1 CO levels
    const mst1Row = { assessment: 'MST1' };
    ['CO1', 'CO2', 'CO3', 'CO4', 'CO5'].forEach(co => {
      // Initialize counters for this CO
      let totalScore = 0;
      let studentCount = 0;
      let studentsAboveTarget = 0;

      // Calculate total score for this CO
      sheets.forEach(student => {
        let coTotal = 0;
        if (coMapping.MST1_Q1 === co && student.MST1_Q1 != null) coTotal += parseFloat(student.MST1_Q1);
        if (coMapping.MST1_Q2 === co && student.MST1_Q2 != null) coTotal += parseFloat(student.MST1_Q2);
        if (coMapping.MST1_Q3 === co && student.MST1_Q3 != null) coTotal += parseFloat(student.MST1_Q3);
        
        if (coTotal > 0) {
          totalScore += coTotal;
          studentCount++;
        }
      });

      if (studentCount > 0) {
        const targetScore = Math.floor(totalScore / studentCount);

        // Count students above target, but only consider students who have valid scores
        let validStudents = 0;
        studentsAboveTarget = 0;
        sheets.forEach(student => {
          let coTotal = 0;
          let hasValidScore = false;
          if (coMapping.MST1_Q1 === co && student.MST1_Q1 != null) {
            coTotal += parseFloat(student.MST1_Q1);
            hasValidScore = true;
          }
          if (coMapping.MST1_Q2 === co && student.MST1_Q2 != null) {
            coTotal += parseFloat(student.MST1_Q2);
            hasValidScore = true;
          }
          if (coMapping.MST1_Q3 === co && student.MST1_Q3 != null) {
            coTotal += parseFloat(student.MST1_Q3);
            hasValidScore = true;
          }
          
          if (hasValidScore) {
            validStudents++;
            if (coTotal >= targetScore) {
              studentsAboveTarget++;
            }
          }
        });

        // Calculate percentage using valid students count
        const percentage = (studentsAboveTarget / validStudents) * 100;
        if (percentage >= 70) mst1Row[co.toLowerCase()] = 3;
        else if (percentage >= 60) mst1Row[co.toLowerCase()] = 2;
        else if (percentage >= 50) mst1Row[co.toLowerCase()] = 1;
        else mst1Row[co.toLowerCase()] = 0;
      } else {
        mst1Row[co.toLowerCase()] = '';
      }
    });

    // Similar logic for MST2
    const mst2Row = { assessment: 'MST2' };
    ['CO1', 'CO2', 'CO3', 'CO4', 'CO5'].forEach(co => {
      let totalScore = 0;
      let studentCount = 0;
      let studentsAboveTarget = 0;

      sheets.forEach(student => {
        let coTotal = 0;
        if (coMapping.MST2_Q1 === co && student.MST2_Q1 != null) coTotal += parseFloat(student.MST2_Q1);
        if (coMapping.MST2_Q2 === co && student.MST2_Q2 != null) coTotal += parseFloat(student.MST2_Q2);
        if (coMapping.MST2_Q3 === co && student.MST2_Q3 != null) coTotal += parseFloat(student.MST2_Q3);
        
        if (coTotal > 0) {
          totalScore += coTotal;
          studentCount++;
        }
      });

      if (studentCount > 0) {
        const targetScore = Math.floor(totalScore / studentCount);

        let validStudents = 0;
        studentsAboveTarget = 0;
        sheets.forEach(student => {
          let coTotal = 0;
          let hasValidScore = false;
          if (coMapping.MST2_Q1 === co && student.MST2_Q1 != null) {
            coTotal += parseFloat(student.MST2_Q1);
            hasValidScore = true;
          }
          if (coMapping.MST2_Q2 === co && student.MST2_Q2 != null) {
            coTotal += parseFloat(student.MST2_Q2);
            hasValidScore = true;
          }
          if (coMapping.MST2_Q3 === co && student.MST2_Q3 != null) {
            coTotal += parseFloat(student.MST2_Q3);
            hasValidScore = true;
          }
          
          if (hasValidScore) {
            validStudents++;
            if (coTotal >= targetScore) {
              studentsAboveTarget++;
            }
          }
        });

        const percentage = (studentsAboveTarget / validStudents) * 100;
        if (percentage >= 70) mst2Row[co.toLowerCase()] = 3;
        else if (percentage >= 60) mst2Row[co.toLowerCase()] = 2;
        else if (percentage >= 50) mst2Row[co.toLowerCase()] = 1;
        else mst2Row[co.toLowerCase()] = 0;
      } else {
        mst2Row[co.toLowerCase()] = '';
      }
    });

    // For Assignment
    const assignmentRow = { assessment: 'Assignment' };
    ['CO1', 'CO2', 'CO3', 'CO4', 'CO5'].forEach(co => {
      let totalScore = 0;
      let studentCount = 0;
      let studentsAboveTarget = 0;

      sheets.forEach(student => {
        const score = student[`Assignment_${co}`];
        if (score != null) {
          totalScore += parseFloat(score);
          studentCount++;
        }
      });

      if (studentCount > 0) {
        const targetScore = Math.floor(totalScore / studentCount);

        let validStudents = 0;
        studentsAboveTarget = 0;
        sheets.forEach(student => {
          const score = student[`Assignment_${co}`];
          if (score != null) {
            validStudents++;
            if (parseFloat(score) >= targetScore) {
              studentsAboveTarget++;
            }
          }
        });

        const percentage = (studentsAboveTarget / validStudents) * 100;
        if (percentage >= 70) assignmentRow[co.toLowerCase()] = 3;
        else if (percentage >= 60) assignmentRow[co.toLowerCase()] = 2;
        else if (percentage >= 50) assignmentRow[co.toLowerCase()] = 1;
        else assignmentRow[co.toLowerCase()] = 0;
      } else {
        assignmentRow[co.toLowerCase()] = '';
      }
    });

    // For EndSem
    const endSemRow = { assessment: 'EndSem' };
    ['CO1', 'CO2', 'CO3', 'CO4', 'CO5'].forEach((co, index) => {
      let totalScore = 0;
      let studentCount = 0;
      let studentsAboveTarget = 0;

      sheets.forEach(student => {
        const score = student[`EndSem_Q${index + 1}`];
        if (score != null) {
          totalScore += parseFloat(score);
          studentCount++;
        }
      });

      if (studentCount > 0) {
        const targetScore = Math.floor(totalScore / studentCount);

        let validStudents = 0;
        studentsAboveTarget = 0;
        sheets.forEach(student => {
          const score = student[`EndSem_Q${index + 1}`];
          if (score != null) {
            validStudents++;
            if (parseFloat(score) >= targetScore) {
              studentsAboveTarget++;
            }
          }
        });

        const percentage = (studentsAboveTarget / validStudents) * 100;
        if (percentage >= 70) endSemRow[co.toLowerCase()] = 3;
        else if (percentage >= 60) endSemRow[co.toLowerCase()] = 2;
        else if (percentage >= 50) endSemRow[co.toLowerCase()] = 1;
        else endSemRow[co.toLowerCase()] = 0;
      } else {
        endSemRow[co.toLowerCase()] = '';
      }
    });

    // Add all rows to worksheet
    worksheet.addRow(mst1Row);
    worksheet.addRow(mst2Row);
    worksheet.addRow(assignmentRow);
    
    // Define and calculate CIE row
    const cieRow = {
      assessment: 'CIE (30%)',
      name: '',
      co1: '',
      co2: '',
      co3: '',
      co4: '',
      co5: ''
    };

    // Calculate CIE values for each CO
    ['co1', 'co2', 'co3', 'co4', 'co5'].forEach(co => {
      const mst1Value = mst1Row[co] || 0;
      const mst2Value = mst2Row[co] || 0;
      const assignmentValue = assignmentRow[co] || 0;
      
      if (mst1Value !== '' || mst2Value !== '' || assignmentValue !== '') {
        cieRow[co] = ((mst1Value + mst2Value + assignmentValue) * 0.3).toFixed(2);
      }
    });

    // Add CIE row
    const cieRowAdded = worksheet.addRow(cieRow);
    cieRowAdded.font = { bold: true };
    cieRowAdded.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE6E6FA' } // Light lavender background
    };

    // Now add EndSem row
    worksheet.addRow(endSemRow);

    // Define and calculate EndSem 70% row
    const endSem70Row = {
      assessment: 'EndSem (70%)',
      name: '',
      co1: '',
      co2: '',
      co3: '',
      co4: '',
      co5: ''
    };

    // Calculate 70% of EndSem values
    ['co1', 'co2', 'co3', 'co4', 'co5'].forEach(co => {
      const endSemValue = endSemRow[co] || 0;
      if (endSemValue !== '') {
        endSem70Row[co] = (endSemValue * 0.7).toFixed(2);
      }
    });

    // Add EndSem 70% row
    const endSem70RowAdded = worksheet.addRow(endSem70Row);
    endSem70RowAdded.font = { bold: true };
    endSem70RowAdded.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFA07A' } // Light salmon background
    };

    // Define and calculate Total row
    const totalRow = {
      assessment: 'Total (CIE + EndSem)',
      name: '',
      co1: '',
      co2: '',
      co3: '',
      co4: '',
      co5: ''
    };

    // Calculate total values
    ['co1', 'co2', 'co3', 'co4', 'co5'].forEach(co => {
      const cieValue = parseFloat(cieRow[co] || 0);
      const endSem70Value = parseFloat(endSem70Row[co] || 0);
      if (cieValue !== 0 || endSem70Value !== 0) {
        totalRow[co] = (cieValue + endSem70Value).toFixed(2);
      }
    });

    // Add Total row
    const totalRowAdded = worksheet.addRow(totalRow);
    totalRowAdded.font = { bold: true };
    totalRowAdded.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF90EE90' } // Light green background
    };

    // Style the rows
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        if (typeof cell.value === 'number') {
          cell.numFmt = '0.00';
        }
      });

      // Add specific colors for different rows
      if (row.getCell(1).value === 'CIE (30%)') {
        row.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFE6E6FA' } // Light lavender
        };
      } else if (row.getCell(1).value === 'EndSem (70%)') {
        row.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFA07A' } // Light salmon
        };
      } else if (row.getCell(1).value === 'Total (CIE + EndSem)') {
        row.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF90EE90' } // Light green
        };
      }
      row.font = { bold: true };
    });

    // Set response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=CO_Matrix_${subjectCode}.xlsx`);

    // Write to response
    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Error generating CO matrix:', error);
    res.status(500).json({ error: 'Failed to generate CO matrix' });
  }
});

// Route to generate and download Excel file for a specific subjectCode
router.get('/end-excel/:subjectCode', async (req, res) => {
  try {
    const { subjectCode } = req.params;

    // Fetch all sheets for the given subject
    const sheets = await prisma.sheet.findMany({
      where: {
        subjectCode: subjectCode
      },
      select: {
        id: true,
        name: true,
        EndSem_Q1: true,
        EndSem_Q2: true,
        EndSem_Q3: true,
        EndSem_Q4: true,
        EndSem_Q5: true,
      }
    });

    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Student Marks');

    // Define headers
    worksheet.columns = [
      { header: 'ID', key: 'id', width: 15 },
      { header: 'Name', key: 'name', width: 20 },
      { header: 'CO1', key: 'co1', width: 10 },
      { header: 'CO2', key: 'co2', width: 10 },
      { header: 'CO3', key: 'co3', width: 10 },
      { header: 'CO4', key: 'co4', width: 10 },
      { header: 'CO5', key: 'co5', width: 10 },
    ];

    // Add data rows
    sheets.forEach(sheet => {
      worksheet.addRow({
        id: sheet.id,
        name: sheet.name,
        co1: sheet.EndSem_Q1 != null ? parseFloat(sheet.EndSem_Q1) : '',
        co2: sheet.EndSem_Q2 != null ? parseFloat(sheet.EndSem_Q2) : '',
        co3: sheet.EndSem_Q3 != null ? parseFloat(sheet.EndSem_Q3) : '',
        co4: sheet.EndSem_Q4 != null ? parseFloat(sheet.EndSem_Q4) : '',
        co5: sheet.EndSem_Q5 != null ? parseFloat(sheet.EndSem_Q5) : '',
      });
    });

    // Calculate target marks (averages)
    const calculateAverage = (values) => {
      const validValues = values.filter(value => value != null);
      const sum = validValues.reduce((acc, value) => acc + parseFloat(value), 0);
      return validValues.length ? Math.floor(sum / validValues.length) : '';
    };

    const targetMarks = {
      id: 'Target Marks',
      name: '',
      co1: calculateAverage(sheets.map(sheet => sheet.EndSem_Q1)),
      co2: calculateAverage(sheets.map(sheet => sheet.EndSem_Q2)),
      co3: calculateAverage(sheets.map(sheet => sheet.EndSem_Q3)),
      co4: calculateAverage(sheets.map(sheet => sheet.EndSem_Q4)),
      co5: calculateAverage(sheets.map(sheet => sheet.EndSem_Q5)),
    };

    // Calculate CO levels
    const calculateCOLevel = (scores, targetMark) => {
      if (targetMark === '') return '';
      const validScores = scores.filter(score => score != null);
      const totalStudents = validScores.length; // Only count students with valid scores
      const studentsAboveTarget = validScores.filter(score => parseFloat(score) >= targetMark).length;
      const percentage = (studentsAboveTarget / totalStudents) * 100;

      if (percentage >= 70) return 3;
      if (percentage >= 60) return 2;
      if (percentage >= 50) return 1;
      return 0;
    };

    const coLevels = {
      id: 'CO Level',
      name: '',
      co1: targetMarks.co1 !== '' ? calculateCOLevel(sheets.map(s => s.EndSem_Q1), targetMarks.co1) : '',
      co2: targetMarks.co2 !== '' ? calculateCOLevel(sheets.map(s => s.EndSem_Q2), targetMarks.co2) : '',
      co3: targetMarks.co3 !== '' ? calculateCOLevel(sheets.map(s => s.EndSem_Q3), targetMarks.co3) : '',
      co4: targetMarks.co4 !== '' ? calculateCOLevel(sheets.map(s => s.EndSem_Q4), targetMarks.co4) : '',
      co5: targetMarks.co5 !== '' ? calculateCOLevel(sheets.map(s => s.EndSem_Q5), targetMarks.co5) : '',
    };

    // Add target marks row
    const targetRow = worksheet.addRow(targetMarks);
    targetRow.font = { bold: true };
    targetRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFD700' } // Light gold background
    };

    // Add students above target row
    const studentsAboveTarget = {
      id: 'Students Above Target',
      name: '',
      co1: targetMarks.co1 !== '' ? sheets.filter(s => s.EndSem_Q1 != null && parseFloat(s.EndSem_Q1) >= targetMarks.co1).length : '',
      co2: targetMarks.co2 !== '' ? sheets.filter(s => s.EndSem_Q2 != null && parseFloat(s.EndSem_Q2) >= targetMarks.co2).length : '',
      co3: targetMarks.co3 !== '' ? sheets.filter(s => s.EndSem_Q3 != null && parseFloat(s.EndSem_Q3) >= targetMarks.co3).length : '',
      co4: targetMarks.co4 !== '' ? sheets.filter(s => s.EndSem_Q4 != null && parseFloat(s.EndSem_Q4) >= targetMarks.co4).length : '',
      co5: targetMarks.co5 !== '' ? sheets.filter(s => s.EndSem_Q5 != null && parseFloat(s.EndSem_Q5) >= targetMarks.co5).length : '',
    };

    // Add students above target row
    const studentsAboveTargetRow = worksheet.addRow(studentsAboveTarget);
    studentsAboveTargetRow.font = { bold: true };
    studentsAboveTargetRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF0E68C' } // Light khaki background
    };
  
    // Add CO level row
    const coLevelRow = worksheet.addRow(coLevels);
    coLevelRow.font = { bold: true };
    coLevelRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF98FB98' } // Light green background
    };

    // Style the headers
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' }
    };

    // Set number format for CO columns to show 2 decimal places
    for (let i = 3; i <= 7; i++) {
      worksheet.getColumn(i).numFmt = '0.00';
    }

    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();

    // Set response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=student_marks_${subjectCode}.xlsx`);

    // Send the file
    res.send(buffer);

  } catch (error) {
    console.error('Error generating Excel file:', error);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});

router.get('/assignment-excel/:subjectCode', async (req, res) => {
  try {
    const { subjectCode } = req.params;

    // Fetch all sheets for the given subject
    const sheets = await prisma.sheet.findMany({
      where: {
        subjectCode: subjectCode
      },
      select: {
        id: true,
        name: true,
        Assignment_CO1: true,
        Assignment_CO2: true,
        Assignment_CO3: true,
        Assignment_CO4: true,
        Assignment_CO5: true,
      }
    });

    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Assignment Marks');

    // Define headers
    worksheet.columns = [
      { header: 'ID', key: 'id', width: 15 },
      { header: 'Name', key: 'name', width: 20 },
      { header: 'CO1', key: 'co1', width: 10 },
      { header: 'CO2', key: 'co2', width: 10 },
      { header: 'CO3', key: 'co3', width: 10 },
      { header: 'CO4', key: 'co4', width: 10 },
      { header: 'CO5', key: 'co5', width: 10 },
    ];

    // Add data rows
    sheets.forEach(sheet => {
      worksheet.addRow({
        id: sheet.id,
        name: sheet.name,
        co1: sheet.Assignment_CO1 != null ? parseFloat(sheet.Assignment_CO1) : '',
        co2: sheet.Assignment_CO2 != null ? parseFloat(sheet.Assignment_CO2) : '',
        co3: sheet.Assignment_CO3 != null ? parseFloat(sheet.Assignment_CO3) : '',
        co4: sheet.Assignment_CO4 != null ? parseFloat(sheet.Assignment_CO4) : '',
        co5: sheet.Assignment_CO5 != null ? parseFloat(sheet.Assignment_CO5) : '',
      });
    });

    // Calculate target marks (averages)
    const totalRows = sheets.length;
    const targetMarks = {
      id: 'Target Marks',
      name: '',
      co1: sheets.some(sheet => sheet.Assignment_CO1 != null) ? 
        Math.floor(sheets.reduce((sum, sheet) => sum + (sheet.Assignment_CO1 != null ? parseFloat(sheet.Assignment_CO1) : 0), 0) / 
        sheets.filter(sheet => sheet.Assignment_CO1 != null).length) : '',
      co2: sheets.some(sheet => sheet.Assignment_CO2 != null) ? 
        Math.floor(sheets.reduce((sum, sheet) => sum + (sheet.Assignment_CO2 != null ? parseFloat(sheet.Assignment_CO2) : 0), 0) / 
        sheets.filter(sheet => sheet.Assignment_CO2 != null).length) : '',
      co3: sheets.some(sheet => sheet.Assignment_CO3 != null) ? 
        Math.floor(sheets.reduce((sum, sheet) => sum + (sheet.Assignment_CO3 != null ? parseFloat(sheet.Assignment_CO3) : 0), 0) / 
        sheets.filter(sheet => sheet.Assignment_CO3 != null).length) : '',
      co4: sheets.some(sheet => sheet.Assignment_CO4 != null) ? 
        Math.floor(sheets.reduce((sum, sheet) => sum + (sheet.Assignment_CO4 != null ? parseFloat(sheet.Assignment_CO4) : 0), 0) / 
        sheets.filter(sheet => sheet.Assignment_CO4 != null).length) : '',
      co5: sheets.some(sheet => sheet.Assignment_CO5 != null) ? 
        Math.floor(sheets.reduce((sum, sheet) => sum + (sheet.Assignment_CO5 != null ? parseFloat(sheet.Assignment_CO5) : 0), 0) / 
        sheets.filter(sheet => sheet.Assignment_CO5 != null).length) : '',
    };

    // Calculate CO levels
    const calculateCOLevel = (scores, targetMark) => {
      if (targetMark === '') return '';
      const validScores = scores.filter(score => score != null);
      const totalStudents = validScores.length; // Only count students with valid scores
      const studentsAboveTarget = validScores.filter(score => parseFloat(score) >= targetMark).length;
      const percentage = (studentsAboveTarget / totalStudents) * 100;

      if (percentage >= 70) return 3;
      if (percentage >= 60) return 2;
      if (percentage >= 50) return 1;
      return 0;
    };

    const coLevels = {
      id: 'CO Level',
      name: '',
      co1: targetMarks.co1 !== '' ? calculateCOLevel(sheets.map(s => s.Assignment_CO1), targetMarks.co1) : '',
      co2: targetMarks.co2 !== '' ? calculateCOLevel(sheets.map(s => s.Assignment_CO2), targetMarks.co2) : '',
      co3: targetMarks.co3 !== '' ? calculateCOLevel(sheets.map(s => s.Assignment_CO3), targetMarks.co3) : '',
      co4: targetMarks.co4 !== '' ? calculateCOLevel(sheets.map(s => s.Assignment_CO4), targetMarks.co4) : '',
      co5: targetMarks.co5 !== '' ? calculateCOLevel(sheets.map(s => s.Assignment_CO5), targetMarks.co5) : '',
    };

    // Add target marks row
    const targetRow = worksheet.addRow(targetMarks);
    targetRow.font = { bold: true };
    targetRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFD700' } // Light gold background
    };

    // Add students above target row
    const studentsAboveTarget = {
      id: 'Students Above Target',
      name: '',
      co1: targetMarks.co1 !== '' ? sheets.filter(s => s.Assignment_CO1 != null && parseFloat(s.Assignment_CO1) >= targetMarks.co1).length : '',
      co2: targetMarks.co2 !== '' ? sheets.filter(s => s.Assignment_CO2 != null && parseFloat(s.Assignment_CO2) >= targetMarks.co2).length : '',
      co3: targetMarks.co3 !== '' ? sheets.filter(s => s.Assignment_CO3 != null && parseFloat(s.Assignment_CO3) >= targetMarks.co3).length : '',
      co4: targetMarks.co4 !== '' ? sheets.filter(s => s.Assignment_CO4 != null && parseFloat(s.Assignment_CO4) >= targetMarks.co4).length : '',
      co5: targetMarks.co5 !== '' ? sheets.filter(s => s.Assignment_CO5 != null && parseFloat(s.Assignment_CO5) >= targetMarks.co5).length : '',
    };

    // Add students above target row
    const studentsAboveTargetRow = worksheet.addRow(studentsAboveTarget);
    studentsAboveTargetRow.font = { bold: true };
    studentsAboveTargetRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF0E68C' } // Light khaki background
    };
  
    // Add CO level row
    const coLevelRow = worksheet.addRow(coLevels);
    coLevelRow.font = { bold: true };
    coLevelRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF98FB98' } // Light green background
    };

    // Style the headers
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' }
    };

    // Set number format for CO columns to show 2 decimal places
    for (let i = 3; i <= 7; i++) {
      worksheet.getColumn(i).numFmt = '0.00';
    }

    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();

    // Set response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=assignment_marks_${subjectCode}.xlsx`);

    // Send the file
    res.send(buffer);

  } catch (error) {
    console.error('Error generating Excel file:', error);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});

// Helper function to calculate MST scores for a specific CO
const calculateMSTScores = (sheets, coMappings) => {
  const scores = {
    MST1: { CO1: 0, CO2: 0, CO3: 0, CO4: 0, CO5: 0 },
    MST2: { CO1: 0, CO2: 0, CO3: 0, CO4: 0, CO5: 0 }
  };
  
  sheets.forEach(sheet => {
    // MST1
    if (sheet.MST1_Q1) {
      scores.MST1[coMappings.MST1_Q1] = (scores.MST1[coMappings.MST1_Q1] || 0) + parseFloat(sheet.MST1_Q1);
    }
    if (sheet.MST1_Q2) {
      scores.MST1[coMappings.MST1_Q2] = (scores.MST1[coMappings.MST1_Q2] || 0) + parseFloat(sheet.MST1_Q2);
    }
    if (sheet.MST1_Q3) {
      scores.MST1[coMappings.MST1_Q3] = (scores.MST1[coMappings.MST1_Q3] || 0) + parseFloat(sheet.MST1_Q3);
    }
    
    // MST2
    if (sheet.MST2_Q1) {
      scores.MST2[coMappings.MST2_Q1] = (scores.MST2[coMappings.MST2_Q1] || 0) + parseFloat(sheet.MST2_Q1);
    }
    if (sheet.MST2_Q2) {
      scores.MST2[coMappings.MST2_Q2] = (scores.MST2[coMappings.MST2_Q2] || 0) + parseFloat(sheet.MST2_Q2);
    }
    if (sheet.MST2_Q3) {
      scores.MST2[coMappings.MST2_Q3] = (scores.MST2[coMappings.MST2_Q3] || 0) + parseFloat(sheet.MST2_Q3);
    }
  });
  
  return scores;
};

// Helper function to calculate Quiz/Assignment scores
const calculateQuizScores = (sheets) => {
  const scores = { CO1: 0, CO2: 0, CO3: 0, CO4: 0, CO5: 0 };
  
  sheets.forEach(sheet => {
    if (sheet.Assignment_CO1) scores.CO1 += parseFloat(sheet.Assignment_CO1);
    if (sheet.Assignment_CO2) scores.CO2 += parseFloat(sheet.Assignment_CO2);
    if (sheet.Assignment_CO3) scores.CO3 += parseFloat(sheet.Assignment_CO3);
    if (sheet.Assignment_CO4) scores.CO4 += parseFloat(sheet.Assignment_CO4);
    if (sheet.Assignment_CO5) scores.CO5 += parseFloat(sheet.Assignment_CO5);
  });
  
  return scores;
};

// Helper function to calculate End Semester scores
const calculateEndSemScores = (sheets) => {
  const scores = {
    CO1: 0,
    CO2: 0,
    CO3: 0,
    CO4: 0,
    CO5: 0
  };
  
  sheets.forEach(sheet => {
    if (sheet.EndSem_Q1) scores.CO1 += parseFloat(sheet.EndSem_Q1);
    if (sheet.EndSem_Q2) scores.CO2 += parseFloat(sheet.EndSem_Q2);
    if (sheet.EndSem_Q3) scores.CO3 += parseFloat(sheet.EndSem_Q3);
    if (sheet.EndSem_Q4) scores.CO4 += parseFloat(sheet.EndSem_Q4);
    if (sheet.EndSem_Q5) scores.CO5 += parseFloat(sheet.EndSem_Q5);
  });
  
  return scores;
};

router.get('/generate-co-attainment/:subjectCode', async (req, res) => {
  try {
    const { subjectCode } = req.params;
    
    // Fetch data for specific subject with correct relations
    const subjectData = await prisma.subject.findUnique({
      where: { code: subjectCode },
      include: {
        sheets: true
      }
    });
    
    if (!subjectData) {
      return res.status(404).json({ error: 'Subject not found' });
    }
    
    // Fetch CO mappings separately
    const coMapping = await prisma.CO.findUnique({
      where: { subjectCode: subjectCode }
    });
    
    if (!coMapping) {
      return res.status(404).json({ error: 'CO mapping not found' });
    }
    
    // Calculate all scores
    const mstScores = calculateMSTScores(subjectData.sheets, coMapping);
    const quizScores = calculateQuizScores(subjectData.sheets, coMapping.Quiz_Assignment);
    const endSemScores = calculateEndSemScores(subjectData.sheets);
    
    const studentCount = subjectData.sheets.length;
    const cos = ['CO1', 'CO2', 'CO3', 'CO4', 'CO5'];
    
    // Calculate final scores
    const finalScores = cos.reduce((acc, co) => {
      const validMST1Students = subjectData.sheets.filter(s => s[`MST1_${co}`] != null).length;
      const validMST2Students = subjectData.sheets.filter(s => s[`MST2_${co}`] != null).length;
      const validQuizStudents = subjectData.sheets.filter(s => s[`Assignment_${co}`] != null).length;
      const validEndSemStudents = subjectData.sheets.filter(s => s[`EndSem_${co}`] != null).length;

      // Calculate CIE (30% weightage) using valid student counts
      const cieScore = (
        ((mstScores.MST1[co] || 0) / (validMST1Students || 1) + 
         (mstScores.MST2[co] || 0) / (validMST2Students || 1) + 
         (quizScores[co] || 0) / (validQuizStudents || 1))
      ) * 0.3;
      
      // Calculate End Sem (70% weightage) using valid student count
      const endSemScore = (endSemScores[co] || 0) / (validEndSemStudents || 1) * 0.7;
      
      acc[co] = {
        MST1: Math.floor((mstScores.MST1[co] || 0) / (validMST1Students || 1)),
        MST2: Math.floor((mstScores.MST2[co] || 0) / (validMST2Students || 1)),
        Quiz: Math.floor((quizScores[co] || 0) / (validQuizStudents || 1)),
        CIE: Math.floor(cieScore),
        EndSem: Math.floor((endSemScores[co] || 0) / (validEndSemStudents || 1)),
        Final: Math.floor(cieScore + endSemScore)
      };
      
      return acc;
    }, {});
    
    // Create Excel workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('CO Attainment');
    
    // Updated column widths to accommodate longer content
    worksheet.columns = [
      { width: 12 },  // CO column
      { width: 20 },  // MST-1
      { width: 20 },  // MST-2
      { width: 25 },  // Assignment/Quiz
      { width: 25 },  // CIE
      { width: 30 },  // End Sem Exam
      { width: 30 },  // CO Direct Attainment
    ];
    
    // Add title
    worksheet.mergeCells('A1:G1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = 'SGSITS Indore - Overall Direct CO attainment';
    titleCell.font = { bold: true, size: 14 };
    titleCell.alignment = { horizontal: 'center' };
    
    // Add headers
    const headers = [
      ['', '', '', '', 'Cumulative Internal Evaluation', 'SEE (Semester End sem Exam)', 'CO Direct Attainment'],
      ['', 'MST-1 (15M)', 'MST-2 (15M)', 'Assignment/Quiz (10)', '(CIE)', '', 'Total'],
      ['', '', '', '', '30% Weightage', '70% Weightage', '30%+70% for each CO']
    ];
    
    worksheet.mergeCells('B3:D3');
    headers.forEach((row, index) => {
      const rowIndex = index + 3;
      worksheet.getRow(rowIndex).values = row;
      worksheet.getRow(rowIndex).font = { bold: true };
    });
    
    // Add data
    cos.forEach((co, index) => {
      const rowIndex = index + 6;
      worksheet.getRow(rowIndex).values = [
        co,
        finalScores[co].MST1,
        finalScores[co].MST2,
        finalScores[co].Quiz,
        finalScores[co].CIE,
        finalScores[co].EndSem,
        finalScores[co].Final
      ];
    });
    
    // Style the data
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        if (typeof cell.value === 'number') {
          cell.numFmt = '0.00';
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
    
    // Set header row background
    ['A3:G3', 'A4:G4', 'A5:G5'].forEach(range => {
      worksheet.getCell(range).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFB4C6E7' }
      };
    });
    
    // Generate response
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=CO_Attainment_${subjectCode}.xlsx`);
    
    await workbook.xlsx.write(res);
    
  } catch (error) {
    console.error('Error generating CO attainment:', error);
    res.status(500).json({ error: 'Failed to generate CO attainment sheet' });
  }
});

module.exports = router;