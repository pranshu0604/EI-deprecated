/*
  Warnings:

  - You are about to drop the column `Quiz_Assignment` on the `Sheet` table. All the data in the column will be lost.

*/
-- AlterTable
ALTER TABLE "Sheet" DROP COLUMN "Quiz_Assignment",
ADD COLUMN     "Quiz_Assignment_CO1" INTEGER,
ADD COLUMN     "Quiz_Assignment_CO2" INTEGER,
ADD COLUMN     "Quiz_Assignment_CO3" INTEGER,
ADD COLUMN     "Quiz_Assignment_CO4" INTEGER,
ADD COLUMN     "Quiz_Assignment_CO5" INTEGER;
