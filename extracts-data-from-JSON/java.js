const fs = require('fs');
const xlsx = require('xlsx');

// Function to extract questions, answers, and feedback from JSON and write to Excel
function extractToExcel(jsonFilePath, excelFilePath) {
    // Read the JSON file
    const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf8'));

    // Prepare data for Excel
    const excelData = [];

    // Iterate through each entry in jsonData
    jsonData.forEach(entry => {
        // Initialize variables to store question, answer, and feedback
        let question = '';
        let answer = '';
        let feedback = '';

        // Iterate through messages in each entry
        entry.messages.forEach(message => {
            if (message.role === 'user') {
                // Extract question from user role message
                question = message.content;
            } else if (message.role === 'bot') {
                // Extract answer from bot role message
                answer = message.content;
                // Extract feedback if available
                feedback = message.feedback || ''; // Use default value if feedback is empty
                
                // Check if the mini conversation is complete
                if (question && answer) {
                    // Push extracted data into excelData array as an object
                    excelData.push({
                        Question: question,
                        Answer: answer,
                        Feedback: feedback
                    });

                    // Reset question and answer for the next mini conversation
                    question = '';
                    answer = '';
                    feedback = '';
                }
            }
        });
    });

    // Create a new workbook and add a worksheet
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(excelData);

    // Append the worksheet to the workbook
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Q&A Data');

    // Write the workbook to an Excel file
    xlsx.writeFile(workbook, excelFilePath);

    console.log(`Data has been written to ${excelFilePath}`);
}

// Example usage
const jsonFilePath = './data.json'; // Path to your JSON file
const excelFilePath = './output.xlsx'; // Path to the output Excel file

extractToExcel(jsonFilePath, excelFilePath);