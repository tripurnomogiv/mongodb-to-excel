# MongoDB to Excel Exporter

A simple Node.js script to export data from MongoDB into an Excel (.xlsx) file.

## Requirements
- Node.js (v16 or later)
- MongoDB (server, local, or Docker)

## Setup

### 1. Install dependencies

npm install

### 2. Configure environment variables

Copy the env example file and adjust the values:

cp .env.example .env

Edit the .env file and update the MongoDB connection settings.

## Run the script

node app.js

## Output

The exported Excel file will be generated automatically inside the output folder.

Example:

output/
└── users.xlsx

## Notes
- The output folder is created automatically if it does not exist.
- .env, node_modules, and exported files are ignored by Git.
