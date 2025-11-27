<%@ page contentType="text/html; charset=UTF-8" %>
<html>
<head>
    <title>Change Password</title>
</head>
<body>

<h2>Change Password</h2>

<!-- Display messages -->
<% if (request.getAttribute("error") != null) { %>
    <p style="color: red;"><%= request.getAttribute("error") %></p>
<% } %>

<% if (request.getAttribute("success") != null) { %>
    <p style="color: green;"><%= request.getAttribute("success") %></p>
<% } %>

<form action="change-password" method="post">
    <input type="password" name="currentPassword" placeholder="Current Password" required><br><br>
    <input type="password" name="newPassword" placeholder="New Password (min 8 chars)" required><br><br>
    <input type="password" name="confirmPassword" placeholder="Confirm New Password" required><br><br>

    <button type="submit">Change Password</button>
</form>

</body>
</html>
