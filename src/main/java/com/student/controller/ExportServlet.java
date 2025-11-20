package com.student.controller;

import com.student.dao.StudentDAO;
import com.student.model.Student;
import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.util.List;

@WebServlet("/export")
public class ExportServlet extends HttpServlet {
    private StudentDAO studentDAO = new StudentDAO();

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        List<Student> students = studentDAO.getAllStudents();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Students");

        int rowCount = 0;

        // Header row
        XSSFRow header = sheet.createRow(rowCount++);
        header.createCell(0).setCellValue("ID");
        header.createCell(1).setCellValue("Student Code");
        header.createCell(2).setCellValue("Full Name");
        header.createCell(3).setCellValue("Email");
        header.createCell(4).setCellValue("Major");

        // Data rows
        for (Student s : students) {
            XSSFRow row = sheet.createRow(rowCount++);
            row.createCell(0).setCellValue(s.getId());
            row.createCell(1).setCellValue(s.getStudentCode());
            row.createCell(2).setCellValue(s.getFullName());
            row.createCell(3).setCellValue(s.getEmail());
            row.createCell(4).setCellValue(s.getMajor());
        }

        // Set response
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=students.xlsx");
        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
