const express = require('express');
const { PrismaClient } = require('@prisma/client');
const router = express.Router();

router.use(express.json());

const prisma = new PrismaClient();

// Route to fetch CO data for a specific subject
router.get('/co-form/:subjectCode', async (req, res) => {
  const { subjectCode } = req.params;

  try {
    const coData = await prisma.cO.findUnique({
      where: { subjectCode },
    });

    if (!coData) {
      return res.status(404).json({ error: 'CO data not found for this subject' });
    }

    res.status(200).json(coData);
  } catch (error) {
    console.error('Error fetching CO data:', error);
    res.status(500).json({ error: 'Failed to fetch CO data' });
  }
});

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
  const { id, name, subjectCode, mst1, mst2, assignment, endsem, indirect } = req.body;

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

  const Indirect_CO1 = indirect?.CO1;
  const Indirect_CO2 = indirect?.CO2;
  const Indirect_CO3 = indirect?.CO3;
  const Indirect_CO4 = indirect?.CO4;
  const Indirect_CO5 = indirect?.CO5;

  try {
    console.log('Received data:', req.body); // Log the received data

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
        MST1_Q1: MST1_Q1 != null ? parseFloat(MST1_Q1) : null,
        MST1_Q2: MST1_Q2 != null ? parseFloat(MST1_Q2) : null,
        MST1_Q3: MST1_Q3 != null ? parseFloat(MST1_Q3) : null,
        MST2_Q1: MST2_Q1 != null ? parseFloat(MST2_Q1) : null,
        MST2_Q2: MST2_Q2 != null ? parseFloat(MST2_Q2) : null,
        MST2_Q3: MST2_Q3 != null ? parseFloat(MST2_Q3) : null,
        EndSem_Q1: EndSem_Q1 != null ? parseFloat(EndSem_Q1) : null,
        EndSem_Q2: EndSem_Q2 != null ? parseFloat(EndSem_Q2) : null,
        EndSem_Q3: EndSem_Q3 != null ? parseFloat(EndSem_Q3) : null,
        EndSem_Q4: EndSem_Q4 != null ? parseFloat(EndSem_Q4) : null,
        EndSem_Q5: EndSem_Q5 != null ? parseFloat(EndSem_Q5) : null,
        Assignment_CO1: Assignment_CO1 != null ? parseFloat(Assignment_CO1) : null,
        Assignment_CO2: Assignment_CO2 != null ? parseFloat(Assignment_CO2) : null,
        Assignment_CO3: Assignment_CO3 != null ? parseFloat(Assignment_CO3) : null,
        Assignment_CO4: Assignment_CO4 != null ? parseFloat(Assignment_CO4) : null,
        Assignment_CO5: Assignment_CO5 != null ? parseFloat(Assignment_CO5) : null,
        Indirect_CO1: Indirect_CO1 != null ? parseFloat(Indirect_CO1) : null,
        Indirect_CO2: Indirect_CO2 != null ? parseFloat(Indirect_CO2) : null,
        Indirect_CO3: Indirect_CO3 != null ? parseFloat(Indirect_CO3) : null,
        Indirect_CO4: Indirect_CO4 != null ? parseFloat(Indirect_CO4) : null,
        Indirect_CO5: Indirect_CO5 != null ? parseFloat(Indirect_CO5) : null,
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
    mst1,
    mst2,
    assignment,
    endsem,
    indirect
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
        MST1_Q1: mst1.Q1,
        MST1_Q2: mst1.Q2,
        MST1_Q3: mst1.Q3,
        MST2_Q1: mst2.Q1,
        MST2_Q2: mst2.Q2,
        MST2_Q3: mst2.Q3,
        Assignment_CO1: assignment.CO1,
        Assignment_CO2: assignment.CO2,
        Assignment_CO3: assignment.CO3,
        Assignment_CO4: assignment.CO4,
        Assignment_CO5: assignment.CO5,
        EndSem_Q1: endsem.Q1,
        EndSem_Q2: endsem.Q2,
        EndSem_Q3: endsem.Q3,
        EndSem_Q4: endsem.Q4,
        EndSem_Q5: endsem.Q5,
        Indirect_CO1: indirect.CO1,
        Indirect_CO2: indirect.CO2,
        Indirect_CO3: indirect.CO3,
        Indirect_CO4: indirect.CO4,
        Indirect_CO5: indirect.CO5
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

// Route to handle Excel file uploads for student data
const multer = require('multer');
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });
const xlsx = require('xlsx');

router.post('/upload-excel/:subjectCode', upload.single('file'), async (req, res) => {
  try {
    const { subjectCode } = req.params;

    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // Read the Excel file buffer
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    if (data.length === 0) {
      return res.status(400).json({ error: 'Excel file contains no data' });
    }

    // Get the subject to find the teacher
    const subject = await prisma.subject.findUnique({
      where: {
        code: subjectCode,
      },
      include: {
        teacher: true,
      },
    });

    if (!subject || !subject.teacher) {
      return res.status(404).json({ error: 'Subject or Teacher not found' });
    }

    // Process each row and create/update student records
    const results = {
      success: 0,
      errors: [],
      updated: 0,
      created: 0
    };

    for (const row of data) {
      try {
        // Check for required fields
        if (!row['Enrollment No.'] || !row['Name']) {
          results.errors.push(`Missing required data for row: ${JSON.stringify(row)}`);
          continue;
        }

        const studentData = {
          id: row['Enrollment No.'].toString(),
          name: row['Name'],
          subjectCode: subjectCode,
          teacherId: subject.teacher.id,

          // MST1 data
          MST1_Q1: row['MST1-Q1'] !== undefined ? parseFloat(row['MST1-Q1']) : null,
          MST1_Q2: row['MST1-Q2'] !== undefined ? parseFloat(row['MST1-Q2']) : null,
          MST1_Q3: row['MST1-Q3'] !== undefined ? parseFloat(row['MST1-Q3']) : null,

          // MST2 data
          MST2_Q1: row['MST2-Q1'] !== undefined ? parseFloat(row['MST2-Q1']) : null,
          MST2_Q2: row['MST2-Q2'] !== undefined ? parseFloat(row['MST2-Q2']) : null,
          MST2_Q3: row['MST2-Q3'] !== undefined ? parseFloat(row['MST2-Q3']) : null,

          // Assignment data
          Assignment_CO1: row['Assignment-CO1'] !== undefined ? parseFloat(row['Assignment-CO1']) : null,
          Assignment_CO2: row['Assignment-CO2'] !== undefined ? parseFloat(row['Assignment-CO2']) : null,
          Assignment_CO3: row['Assignment-CO3'] !== undefined ? parseFloat(row['Assignment-CO3']) : null,
          Assignment_CO4: row['Assignment-CO4'] !== undefined ? parseFloat(row['Assignment-CO4']) : null,
          Assignment_CO5: row['Assignment-CO5'] !== undefined ? parseFloat(row['Assignment-CO5']) : null,

          // EndSem data
          EndSem_Q1: row['EndSem-Q1'] !== undefined ? parseFloat(row['EndSem-Q1']) : null,
          EndSem_Q2: row['EndSem-Q2'] !== undefined ? parseFloat(row['EndSem-Q2']) : null,
          EndSem_Q3: row['EndSem-Q3'] !== undefined ? parseFloat(row['EndSem-Q3']) : null,
          EndSem_Q4: row['EndSem-Q4'] !== undefined ? parseFloat(row['EndSem-Q4']) : null,
          EndSem_Q5: row['EndSem-Q5'] !== undefined ? parseFloat(row['EndSem-Q5']) : null,

          // Indirect CO data
          Indirect_CO1: row['Indirect-CO1'] !== undefined ? parseFloat(row['Indirect-CO1']) : null,
          Indirect_CO2: row['Indirect-CO2'] !== undefined ? parseFloat(row['Indirect-CO2']) : null,
          Indirect_CO3: row['Indirect-CO3'] !== undefined ? parseFloat(row['Indirect-CO3']) : null,
          Indirect_CO4: row['Indirect-CO4'] !== undefined ? parseFloat(row['Indirect-CO4']) : null,
          Indirect_CO5: row['Indirect-CO5'] !== undefined ? parseFloat(row['Indirect-CO5']) : null,
        };

        // Check if student already exists for this subject
        const existingStudent = await prisma.sheet.findUnique({
          where: {
            id_subjectCode: {
              id: studentData.id,
              subjectCode: subjectCode
            }
          }
        });

        if (existingStudent) {
          // Update existing student
          await prisma.sheet.update({
            where: {
              id_subjectCode: {
                id: studentData.id,
                subjectCode: subjectCode
              }
            },
            data: studentData
          });
          results.updated++;
        } else {
          // Create new student record
          await prisma.sheet.create({
            data: studentData
          });
          results.created++;
        }

        results.success++;
      } catch (error) {
        results.errors.push(`Error processing row: ${error.message}`);
      }
    }

    res.status(200).json({
      message: 'Excel data processed successfully',
      results
    });
  } catch (error) {
    console.error('Error processing Excel upload:', error);
    res.status(500).json({ error: 'Failed to process Excel file' });
  }
});

// Route to provide an Excel template for data import
router.get('/excel-template/:subjectCode', async (req, res) => {
  try {
    const { subjectCode } = req.params;
    
    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Student Data Template');
    
    // Set up headers
    worksheet.columns = [
      { header: 'Enrollment No.', key: 'enrollment', width: 15 },
      { header: 'Name', key: 'name', width: 20 },
      { header: 'MST1-Q1', key: 'mst1q1', width: 10 },
      { header: 'MST1-Q2', key: 'mst1q2', width: 10 },
      { header: 'MST1-Q3', key: 'mst1q3', width: 10 },
      { header: 'MST2-Q1', key: 'mst2q1', width: 10 },
      { header: 'MST2-Q2', key: 'mst2q2', width: 10 },
      { header: 'MST2-Q3', key: 'mst2q3', width: 10 },
      { header: 'Assignment-CO1', key: 'assignco1', width: 15 },
      { header: 'Assignment-CO2', key: 'assignco2', width: 15 },
      { header: 'Assignment-CO3', key: 'assignco3', width: 15 },
      { header: 'Assignment-CO4', key: 'assignco4', width: 15 },
      { header: 'Assignment-CO5', key: 'assignco5', width: 15 },
      { header: 'EndSem-Q1', key: 'endsemq1', width: 10 },
      { header: 'EndSem-Q2', key: 'endsemq2', width: 10 },
      { header: 'EndSem-Q3', key: 'endsemq3', width: 10 },
      { header: 'EndSem-Q4', key: 'endsemq4', width: 10 },
      { header: 'EndSem-Q5', key: 'endsemq5', width: 10 },
      { header: 'Indirect-CO1', key: 'indirectco1', width: 12 },
      { header: 'Indirect-CO2', key: 'indirectco2', width: 12 },
      { header: 'Indirect-CO3', key: 'indirectco3', width: 12 },
      { header: 'Indirect-CO4', key: 'indirectco4', width: 12 },
      { header: 'Indirect-CO5', key: 'indirectco5', width: 12 }
    ];
    
    // Try to get existing data for the subject
    try {
      const sheets = await prisma.sheet.findMany({
        where: { subjectCode }
      });
      
      // Add sample data if available
      sheets.forEach(sheet => {
        worksheet.addRow({
          enrollment: sheet.id,
          name: sheet.name,
          mst1q1: sheet.MST1_Q1,
          mst1q2: sheet.MST1_Q2,
          mst1q3: sheet.MST1_Q3,
          mst2q1: sheet.MST2_Q1,
          mst2q2: sheet.MST2_Q2,
          mst2q3: sheet.MST2_Q3,
          assignco1: sheet.Assignment_CO1,
          assignco2: sheet.Assignment_CO2,
          assignco3: sheet.Assignment_CO3,
          assignco4: sheet.Assignment_CO4,
          assignco5: sheet.Assignment_CO5,
          endsemq1: sheet.EndSem_Q1,
          endsemq2: sheet.EndSem_Q2,
          endsemq3: sheet.EndSem_Q3,
          endsemq4: sheet.EndSem_Q4,
          endsemq5: sheet.EndSem_Q5,
          indirectco1: sheet.Indirect_CO1,
          indirectco2: sheet.Indirect_CO2,
          indirectco3: sheet.Indirect_CO3,
          indirectco4: sheet.Indirect_CO4,
          indirectco5: sheet.Indirect_CO5
        });
      });
    } catch (error) {
      console.error("Error fetching existing data:", error);
      // Add a sample empty row 
      worksheet.addRow({
        enrollment: "12345678",
        name: "John Doe",
        mst1q1: 4,
        mst1q2: 5,
        mst1q3: 3,
        mst2q1: 4,
        mst2q2: 4,
        mst2q3: 5,
        assignco1: 8,
        assignco2: 7,
        assignco3: 9,
        assignco4: 8,
        assignco5: 7,
        endsemq1: 10,
        endsemq2: 12,
        endsemq3: 11,
        endsemq4: 13,
        endsemq5: 9,
        indirectco1: 2,
        indirectco2: 3,
        indirectco3: 2,
        indirectco4: 3,
        indirectco5: 2
      });
    }
    
    // Style the header row
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' }
    };
    
    // Set response headers for file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=student_data_template_${subjectCode}.xlsx`);

    // Write the file to the response
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error generating Excel template:', error);
    res.status(500).json({ error: 'Failed to generate Excel template' });
  }
});

module.exports = router;