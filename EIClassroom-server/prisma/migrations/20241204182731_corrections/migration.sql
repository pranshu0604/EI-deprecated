/*
  Warnings:

  - You are about to drop the column `Quiz_Assignment_CO1` on the `Sheet` table. All the data in the column will be lost.
  - You are about to drop the column `Quiz_Assignment_CO2` on the `Sheet` table. All the data in the column will be lost.
  - You are about to drop the column `Quiz_Assignment_CO3` on the `Sheet` table. All the data in the column will be lost.
  - You are about to drop the column `Quiz_Assignment_CO4` on the `Sheet` table. All the data in the column will be lost.
  - You are about to drop the column `Quiz_Assignment_CO5` on the `Sheet` table. All the data in the column will be lost.

*/
-- AlterTable
ALTER TABLE "Sheet" DROP COLUMN "Quiz_Assignment_CO1",
DROP COLUMN "Quiz_Assignment_CO2",
DROP COLUMN "Quiz_Assignment_CO3",
DROP COLUMN "Quiz_Assignment_CO4",
DROP COLUMN "Quiz_Assignment_CO5",
ADD COLUMN     "Assignment_CO1" INTEGER,
ADD COLUMN     "Assignment_CO2" INTEGER,
ADD COLUMN     "Assignment_CO3" INTEGER,
ADD COLUMN     "Assignment_CO4" INTEGER,
ADD COLUMN     "Assignment_CO5" INTEGER;
