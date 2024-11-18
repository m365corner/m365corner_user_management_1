
// Initialize MSAL Configuration
const msalConfig = {
    auth: {
        clientId: "<your-app-id-or-client-id-goes-here>", // Replace with your Azure AD App's Client ID
        authority: "https://login.microsoftonline.com/<your-tenant-id-goes-here>", // Replace with Tenant ID
        redirectUri: "http://localhost:8000", // Replace with your Redirect URI, If using localhost, you can use the same value. But ensure you include this while registering your app. 
    },
    cache: {
        cacheLocation: "localStorage", // Stores tokens in localStorage
        storeAuthStateInCookie: false, // Set true for older browsers
    },
};

// Create MSAL instance
let msalInstance;
try {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    console.log("MSAL Instance initialized successfully.");
} catch (error) {
    console.error("Error initializing MSAL instance:", error);
}

// Acquire token silently or fallback to popup login
async function acquireToken(scopes) {
    const account = msalInstance.getAllAccounts()[0];
    if (!account) {
        throw new Error("No account found. Please sign in first.");
    }

    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes,
            account,
        });
        return tokenResponse.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            return await msalInstance.acquireTokenPopup({ scopes });
        } else {
            throw error;
        }
    }
}

// Login
async function login() {
    if (!msalInstance) {
        console.error("MSAL instance is not initialized!");
        document.getElementById("output").innerText = "Login failed: MSAL instance is not initialized.";
        return;
    }

    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["Mail.Send", "User.ReadWrite.All", "Directory.ReadWrite.All"], // Required scopes
        });
        console.log("Login successful:", loginResponse);
        msalInstance.setActiveAccount(loginResponse.account);
        document.getElementById("output").innerText = "Login successful!";
    } catch (error) {
        console.error("Login failed:", error);
        document.getElementById("output").innerText = `Login failed: ${error}`;
    }
}

// Logout
function logout() {
    if (!msalInstance) {
        console.error("MSAL instance is not initialized!");
        document.getElementById("output").innerText = "Logout failed: MSAL instance is not initialized.";
        return;
    }

    msalInstance.logoutPopup();
    document.getElementById("output").innerText = "Logged out.";
}




// Graph API helper

// Graph API Helper Function
async function callGraphApi(endpoint, method = "GET", body = null) {
    try {
        const token = await acquireToken(["Mail.Send", "User.ReadWrite.All", "Directory.ReadWrite.All"]);
        const headers = new Headers({
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
        });

        const options = {
            method,
            headers,
        };

        if (body) options.body = JSON.stringify(body);

        const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, options);

        if (response.ok) {
            // Safely handle JSON responses
            const contentType = response.headers.get("content-type");
            return contentType && contentType.includes("application/json") ? await response.json() : {};
        }

        // Handle errors gracefully
        const errorContentType = response.headers.get("content-type");
        const errorResponse = errorContentType && errorContentType.includes("application/json")
            ? await response.json()
            : await response.text();
        console.error(`Graph API Error: ${errorResponse}`);
        throw new Error(`Graph API call failed with status ${response.status}`);
    } catch (error) {
        console.error("Error in callGraphApi:", error.message);
        throw error;
    }
}

// Send Report as Mail
async function sendReportAsMail() {
    const recipientEmail = document.getElementById("recipientEmail").value;

    if (!recipientEmail) {
        alert("Please enter a valid recipient email.");
        return;
    }

    // Extract data from the table
    const tableHeaders = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const tableRows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (tableRows.length === 0) {
        alert("No data to send. Please retrieve and display user details first.");
        return;
    }

    // Format the email body as an HTML table
    const emailTable = `
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <thead>
                <tr>${tableHeaders.map(header => `<th>${header}</th>`).join("")}</tr>
            </thead>
            <tbody>
                ${tableRows
                    .map(
                        row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`
                    )
                    .join("")}
            </tbody>
        </table>
    `;

    // Email content
    const email = {
        message: {
            subject: "User Report from M365 User Management Tool",
            body: {
                contentType: "HTML",
                content: `
                    <p>Dear Administrator,</p>
                    <p>Please find below the user report generated by the M365 User Management Tool:</p>
                    ${emailTable}
                    <p>Regards,<br>M365 User Management Team</p>
                `
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: recipientEmail
                    }
                }
            ]
        }
    };

    try {
        const response = await callGraphApi("/me/sendMail", "POST", email);
        alert("Report sent successfully!");
        console.log("Mail Response:", response);
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report. Please try again.");
    }
}

// Acquire Token Helper Function
async function acquireToken(scopes) {
    const account = msalInstance.getActiveAccount();
    if (!account) {
        throw new Error("No account found. Please sign in first.");
    }

    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes,
            account,
        });
        return tokenResponse.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            return await msalInstance.acquireTokenPopup({ scopes });
        } else {
            throw error;
        }
    }
}











// Populate table with JSON data
function populateTable(data) {
    const headers = Object.keys(data[0] || {});
    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");

    // Clear table
    outputHeader.innerHTML = "";
    outputBody.innerHTML = "";

    // Create table headers
    headers.forEach((header) => {
        const th = document.createElement("th");
        th.textContent = header;
        outputHeader.appendChild(th);
    });

    // Populate rows
    data.forEach((row) => {
        const tr = document.createElement("tr");
        headers.forEach((header) => {
            const td = document.createElement("td");
            td.textContent = row[header] || "";
            tr.appendChild(td);
        });
        outputBody.appendChild(tr);
    });
}

// Retrieve user details
async function retrieveUserDetails() {
    try {
        const response = await callGraphApi("/users?$filter=accountEnabled eq true");
        //populateTable(response.value);
        populatePaginatedData(response.value); 
    } catch (error) {
        console.error("Error retrieving user details:", error);
    }
}










// Filter users by department
async function filterUsersByDepartment() {
    const department = document.getElementById("departmentFilter").value;

    if (!department) {
        alert("Please select a department!");
        return;
    }

    try {
        // Build OData query for department filter
        const query = `/users?$filter=department eq '${department}'&$select=id,displayName,userPrincipalName,department`;
        const response = await callGraphApi(query);

        if (response.value && response.value.length > 0) {
            populateTable(response.value);
        } else {
            alert("No users found in the selected department.");
            document.getElementById("outputHeader").innerHTML = "";
            document.getElementById("outputBody").innerHTML = "";
        }
    } catch (error) {
        console.error("Error filtering users by department:", error);
        alert("An error occurred while fetching users for the selected department.");
    }
}


// Filter users by job title
async function filterUsersByJobTitle() {
    const jobTitle = document.getElementById("jobTitleFilter").value;

    if (!jobTitle) {
        alert("Please select a job title!");
        return;
    }

    try {
        // Build OData query for job title filter
        const query = `/users?$filter=jobTitle eq '${jobTitle}'&$select=id,displayName,userPrincipalName,jobTitle`;
        const response = await callGraphApi(query);

        if (response.value && response.value.length > 0) {
            populateTable(response.value);
        } else {
            alert("No users found with the selected job title.");
            document.getElementById("outputHeader").innerHTML = "";
            document.getElementById("outputBody").innerHTML = "";
        }
    } catch (error) {
        console.error("Error filtering users by job title:", error);
        alert("An error occurred while fetching users for the selected job title.");
    }
}


// Filter users by admin role
// Filter users by admin role
// Filter users by admin role
async function filterUsersByAdminRole() {
    const role = document.getElementById("adminRoleFilter").value;

    if (!role) {
        alert("Please select an admin role!");
        return;
    }

    try {
        // Map roles to directory role IDs
        const roleIds = {
            "Global Administrator": "232142d7-3931-4598-b199-75199c53beb7",
            "Security Administrator": "0b847090-daa8-4402-9095-6805f96ba602",
            "Exchange Administrator": "86a973a7-7eb2-478a-aec6-ed52099dc61e",
            "Application Administrator": "22ba8ddb-3a88-43ff-8cf9-8bf8a0d39ad6",
            "Helpdesk Administrator": "7781c44e-27a1-4954-8c90-f257acd68cac",
            "User Administrator": "d325a324-b5e9-4411-b0cf-2861f1333650",
            "Reports Reader": "e73eb654-8a32-43fc-99a1-42313cc284fc",
            "Teams Administrator": "2672bf14-e16c-46ce-9045-3ca16e77b658",
            "SharePoint Administrator": "81067de9-0f58-44be-a313-981b941ad7c9"
        };

        const roleId = roleIds[role];
        if (!roleId) {
            alert("Invalid role selected.");
            return;
        }

        console.log(`Fetching members for role: ${role} (Role ID: ${roleId})`);

        // Call the Graph API
        const query = `/directoryRoles/${roleId}/members?$select=id,displayName,userPrincipalName`;
        console.log(`Query: ${query}`);
        const response = await callGraphApi(query);

        console.log("Graph API response:", response);

        // Handle results
        if (response && response.value && response.value.length > 0) {
            populateTable(response.value);
            console.log(`Populated table with ${response.value.length} users.`);
        } else {
            alert(`No users found with the selected admin role: ${role}.`);
            document.getElementById("outputHeader").innerHTML = "";
            document.getElementById("outputBody").innerHTML = "";
        }
    } catch (error) {
        console.error("Error filtering users by admin role:", error);

        // Debug specific Graph API errors
        if (error.message.includes("Request_ResourceNotFound")) {
            alert("Error: The selected admin role does not exist in the tenant.");
        } else if (error.message.includes("AccessDenied")) {
            alert("Error: You do not have permission to access this resource.");
        } else {
            alert("An unexpected error occurred. Please check the console for more details.");
        }
    }
}


// Filter users by location

async function filterUsersByLocation() {
    const location = document.getElementById("locationFilter").value;

    // Map dropdown values to ISO country codes
    const locationMapping = {
        "United States": "US",
        "China": "CN",
        "India": "IN", 
        "Barbados": "BB"
    };

    const usageLocation = locationMapping[location];

    if (!usageLocation) {
        alert("Please select a valid location!");
        return;
    }

    try {
        // Build OData query for location filter
        const query = `/users?$filter=usageLocation eq '${usageLocation}'&$select=id,displayName,userPrincipalName,usageLocation`;
        console.log(`Query: ${query}`); // Debugging log

        const response = await callGraphApi(query);
        console.log("Graph API response:", response); // Debugging log

        if (response.value && response.value.length > 0) {
            populateTable(response.value);
            console.log(`Populated table with ${response.value.length} users.`);
        } else {
            alert(`No users found in the selected location: ${location}.`);
            document.getElementById("outputHeader").innerHTML = "";
            document.getElementById("outputBody").innerHTML = "";
        }
    } catch (error) {
        console.error("Error filtering users by location:", error);

        if (error.message.includes("AccessDenied")) {
            alert("Error: You do not have permission to access this resource.");
        } else {
            alert("An unexpected error occurred. Please check the console for more details.");
        }
    }
}


// Force a new token acquisition
async function acquireFreshToken(scopes) {
    try {
        console.log("Forcing a new token acquisition...");
        
        // Log out to clear the cache
        await msalInstance.logoutPopup();

        // Log back in and acquire a fresh token
        const loginResponse = await msalInstance.loginPopup({ scopes });
        console.log("Login response:", loginResponse);

        const tokenResponse = await msalInstance.acquireTokenPopup({ scopes });
        console.log("Acquired new token with scopes:", tokenResponse.scopes);
        return tokenResponse.accessToken;
    } catch (error) {
        console.error("Error acquiring a new token:", error);
        throw new Error("Failed to acquire a new token. Please try again.");
    }
}



// Send Report as Mail
async function sendReportAsMail() {
    const recipientEmail = document.getElementById("recipientEmail").value;

    if (!recipientEmail) {
        alert("Please enter a valid recipient email.");
        return;
    }

    // Extract data from the table
    const tableHeaders = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const tableRows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (tableRows.length === 0) {
        alert("No data to send. Please retrieve and display user details first.");
        return;
    }

    // Format the email body as an HTML table
    const emailTable = `
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <thead>
                <tr>${tableHeaders.map(header => `<th>${header}</th>`).join("")}</tr>
            </thead>
            <tbody>
                ${tableRows
                    .map(
                        row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`
                    )
                    .join("")}
            </tbody>
        </table>
    `;

    // Email content
    const email = {
        message: {
            subject: "User Report from M365 User Management Tool",
            body: {
                contentType: "HTML",
                content: `
                    <p>Dear Administrator,</p>
                    <p>Please find below the user report generated by the M365 User Management Tool:</p>
                    ${emailTable}
                    <p>Regards,<br>M365 User Management Team</p>
                `
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: recipientEmail
                    }
                }
            ]
        }
    };

    try {
        const response = await callGraphApi("/me/sendMail", "POST", email);
        alert("Report sent successfully!");
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report. Please try again.");
    }
}


// Search Users by Multiple Properties
async function searchUsers() {
    const query = document.getElementById("searchInput").value.trim();

    if (!query) {
        alert("Please enter a search query.");
        return;
    }

    try {
        // Build OData filter query
        const filterQuery = `
            startswith(givenName,'${query}') or
            startswith(surname,'${query}') or
            startswith(userPrincipalName,'${query}') or
            startswith(displayName,'${query}') or
            startswith(mail,'${query}')
        `.trim();

        const encodedFilter = encodeURIComponent(filterQuery);
        const response = await callGraphApi(`/users?$filter=${encodedFilter}&$select=id,displayName,userPrincipalName,mail,givenName,surname`);

        if (response.value && response.value.length > 0) {
            populateSearchResults(response.value);
        } else {
            alert("No users found for the given search query.");
            clearSearchResults();
        }
    } catch (error) {
        console.error("Error searching users:", error);
        alert("An error occurred while searching for users. Please try again.");
    }
}

// Populate Search Results
function populateSearchResults(users) {
    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");

    if (!outputHeader || !outputBody) {
        console.error("Error: Table elements not found in DOM.");
        return;
    }

    // Set table headers
    outputHeader.innerHTML = `
        <th>First Name</th>
        <th>Last Name</th>
        <th>UserPrincipalName</th>
        <th>Display Name</th>
        <th>Email Address</th>
    `;

    // Populate table body with user data
    outputBody.innerHTML = users
        .map(user => `
            <tr>
                <td>${user.givenName || ""}</td>
                <td>${user.surname || ""}</td>
                <td>${user.userPrincipalName || ""}</td>
                <td>${user.displayName || ""}</td>
                <td>${user.mail || ""}</td>
            </tr>
        `)
        .join("");
}

// Clear Search Results
function clearSearchResults() {
    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");

    if (outputHeader) {
        outputHeader.innerHTML = "";
    }
    if (outputBody) {
        outputBody.innerHTML = "";
    }
}


// Download the user report as a CSV file
function downloadReportAsCSV() {
    const tableHeaders = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const tableRows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (tableRows.length === 0) {
        alert("No data to download. Please retrieve and display user details first.");
        return;
    }

    // Construct CSV content
    let csvContent = tableHeaders.join(",") + "\n"; // Add headers
    csvContent += tableRows.map(row => row.join(",")).join("\n"); // Add rows

    // Create a downloadable link for the CSV
    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "User_Report.csv";
    a.click();

    // Clean up
    URL.revokeObjectURL(url);
}




let currentPage = 1; // Track the current page
const recordsPerPage = 10; // Number of records per page
let paginatedData = []; // Stores paginated data for navigation

// Update pagination controls
function updatePaginationControls() {
    const totalPages = Math.ceil(paginatedData.length / recordsPerPage);
    document.getElementById("pageInfo").innerText = `Page ${currentPage} of ${totalPages}`;
    document.getElementById("prevButton").disabled = currentPage === 1;
    document.getElementById("nextButton").disabled = currentPage === totalPages;
}

// Display the current page
function displayCurrentPage() {
    const startIndex = (currentPage - 1) * recordsPerPage;
    const endIndex = startIndex + recordsPerPage;
    const currentData = paginatedData.slice(startIndex, endIndex);

    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");

    // Set headers if not already populated
    if (outputHeader.innerHTML.trim() === "") {
        outputHeader.innerHTML = `
            <th>First Name</th>
            <th>Last Name</th>
            <th>UserPrincipalName</th>
            <th>Display Name</th>
            <th>Email Address</th>
        `;
    }

    // Populate the table body
    outputBody.innerHTML = currentData
        .map(user => `
            <tr>
                <td>${user.givenName || ""}</td>
                <td>${user.surname || ""}</td>
                <td>${user.userPrincipalName || ""}</td>
                <td>${user.displayName || ""}</td>
                <td>${user.mail || ""}</td>
            </tr>
        `)
        .join("");

    updatePaginationControls();
}

// Handle "Next" button click
function nextPage() {
    if (currentPage * recordsPerPage < paginatedData.length) {
        currentPage++;
        displayCurrentPage();
    }
}

// Handle "Previous" button click
function prevPage() {
    if (currentPage > 1) {
        currentPage--;
        displayCurrentPage();
    }
}

// Populate paginated data and display the first page
function populatePaginatedData(users) {
    paginatedData = users; // Save data for navigation
    currentPage = 1; // Reset to the first page
    displayCurrentPage();
}



// Reset Screen Functionality
function resetScreen() {
    // Clear all filters and input fields
    document.getElementById("searchInput").value = "";
    document.getElementById("recipientEmail").value = "";
    document.getElementById("departmentFilter").value = "";
    document.getElementById("jobTitleFilter").value = "";
    document.getElementById("adminRoleFilter").value = "";
    document.getElementById("locationFilter").value = "";

    // Clear table data
    document.getElementById("outputHeader").innerHTML = "";
    document.getElementById("outputBody").innerHTML = "";

    // Reset pagination
    currentPage = 1;
    paginatedData = [];
    updatePaginationControls();

    alert("Screen has been reset.");
}



