const express = require('express');
const { PrismaClient } = require('@prisma/client');
const authenticateToken = require('../middleware/auth');
const prisma = new PrismaClient();
const router = express.Router();


// Create a new subject
router.post('/newsubject', authenticateToken, async (req, res) => {
  const { name, code } = req.body;
  const teacherId = req.teacher.teacherId;

  try {
    const subject = await prisma.subject.create({
      data: { name, code, teacherId },
    });
    res.status(201).json(subject);
  } catch (error) {
    console.error("Error creating subject:", error);
    res.status(500).json({ error: "Failed to create subject" });
  }
});

//get all subjects for a teacher

router.get('/subjects', authenticateToken, async (req, res) => {
  const teacherId = req.teacher.teacherId;

  try {
    const subjects = await prisma.subject.findMany({
      where: { teacherId },
    });
    res.json(subjects);
  } catch (error) {
    console.error("Error fetching subjects:", error);
    res.status(500).json({ error: "Failed to fetch subjects" });
  }
});
//get all subjects irrespective of teacher
router.get('/allsubjects', async (req, res) => {
  try {
    const subjects = await prisma.subject.findMany();
    res.json(subjects);
  } catch (error) {
    console.error("Error fetching subjects:", error);
    res.status(500).json({ error: "Failed to fetch subjects" });
  }
});

module.exports = router;