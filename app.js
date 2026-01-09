require("dotenv").config();
const { MongoClient, ObjectId } = require("mongodb");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

async function exportMongoToExcel() {
  const DEFAULT_MONGO_URI = "mongodb://127.0.0.1:27017";
  const uri = process.env.MONGO_URI || DEFAULT_MONGO_URI;

  console.log("MongoDB URI:", uri); 

  const client = new MongoClient(uri, {
    maxPoolSize: 10,
    serverSelectionTimeoutMS: 5000
  });

  try {
    // Connect to MongoDB
    await client.connect();
    console.log("MongoDB connected");

    const db = client.db("dbname");
    const collection = db.collection("user");

    // Filter query
    const filter = {
      companyID: new ObjectId("64c486a30ea84b0d15549d8d"),
      active: false,
      //locationID: { $exists: true }
    };

    // Fetch data
    const data = await collection.find(filter).toArray();
    console.log("Total data:", data.length);

    if (data.length === 0) {
      console.log("no data to export");
      return;
    }

    // Create Excel file
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Users");

    // Auto-generate headers from object keys
    worksheet.columns = Object.keys(data[0]).map(key => ({
      header: key,
      key: key,
      width: 25
    }));

    // Insert rows
    data.forEach(item => {
      worksheet.addRow(item);
    });

    // Save Excel file
    const outputPath = path.join(outputDir, "users.xlsx");
    await workbook.xlsx.writeFile(outputPath);
    console.log("Excel file successfully created");

  } catch (error) {
    console.error("Export error:", error);
  } finally {
    // Ensure MongoDB connection is closed
    await client.close();
    console.log("MongoDB connection closed");
  }
}

exportMongoToExcel();