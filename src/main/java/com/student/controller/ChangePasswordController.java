/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.student.controller;

import com.student.dao.UserDAO;
import model.User;

import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;

import java.io.IOException;
import java.security.MessageDigest;

@WebServlet("/change-password")
public class ChangePasswordController extends HttpServlet {

    private UserDAO userDAO;

    @Override
    public void init() throws ServletException {
        userDAO = new UserDAO();
    }

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp)
            throws ServletException, IOException {

        req.getRequestDispatcher("views/change-password.jsp").forward(req, resp);
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp)
            throws ServletException, IOException {

        HttpSession session = req.getSession(false);

        if (session == null || session.getAttribute("user") == null) {
            req.setAttribute("error", "You must be logged in.");
            req.getRequestDispatcher("views/login.jsp").forward(req, resp);
            return;
        }

        User user = (User) session.getAttribute("user");

        String currentPassword = req.getParameter("currentPassword");
        String newPassword = req.getParameter("newPassword");
        String confirmPassword = req.getParameter("confirmPassword");

        String hashedCurrent = hashPassword(currentPassword);

        if (!hashedCurrent.equals(user.getPassword())) {
            req.setAttribute("error", "Current password is incorrect.");
            req.getRequestDispatcher("views/change-password.jsp").forward(req, resp);
            return;
        }

        if (newPassword == null || newPassword.length() < 8) {
            req.setAttribute("error", "New password must be at least 8 characters.");
            req.getRequestDispatcher("views/change-password.jsp").forward(req, resp);
            return;
        }

        if (!newPassword.equals(confirmPassword)) {
            req.setAttribute("error", "New password and confirmation do not match.");
            req.getRequestDispatcher("views/change-password.jsp").forward(req, resp);
            return;
        }

        String newHashed = hashPassword(newPassword);
        boolean updated = userDAO.updatePassword(user.getId(), newHashed);

        if (updated) {
            user.setPassword(newHashed);
            req.setAttribute("success", "Password changed successfully!");
        } else {
            req.setAttribute("error", "Failed to update password.");
        }

        req.getRequestDispatcher("views/change-password.jsp").forward(req, resp);
    }

    private String hashPassword(String password) {
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            md.update(password.getBytes());
            byte[] bytes = md.digest();

            StringBuilder sb = new StringBuilder();
            for (byte b : bytes) sb.append(String.format("%02x", b));
            return sb.toString();

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
