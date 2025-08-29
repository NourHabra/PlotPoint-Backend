/*
  Script: Create default admin user
  Usage: node backend/scripts/create-admin.js
  Env:   MONGODB_URI (optional) - defaults to mongodb://localhost:27017/template-db
*/

const mongoose = require("mongoose");
const bcrypt = require("bcryptjs");
const path = require("path");
require("dotenv").config({ path: path.join(__dirname, "..", ".env") });

const User = require("../models/User");

const ADMIN_NAME = "ReportGen Admin";
const ADMIN_EMAIL = "admin@reportgen.com";
const ADMIN_PASSWORD = "report123";

async function main() {
	const uri =
		process.env.MONGODB_URI || "mongodb://localhost:27017/template-db";
	console.log(`[create-admin] Connecting to ${uri}`);
	await mongoose.connect(uri, {
		useNewUrlParser: true,
		useUnifiedTopology: true,
	});

	try {
		const salt = await bcrypt.genSalt(10);
		const passwordHash = await bcrypt.hash(ADMIN_PASSWORD, salt);

		const existing = await User.findOne({ email: ADMIN_EMAIL });
		if (existing) {
			existing.name = ADMIN_NAME;
			existing.role = "Admin";
			existing.passwordHash = passwordHash;
			await existing.save();
			console.log(
				`[create-admin] Updated existing admin: ${ADMIN_EMAIL}`
			);
		} else {
			await User.create({
				name: ADMIN_NAME,
				email: ADMIN_EMAIL,
				passwordHash,
				role: "Admin",
				avatarPath: "",
			});
			console.log(`[create-admin] Created admin: ${ADMIN_EMAIL}`);
		}
	} catch (err) {
		console.error(
			"[create-admin] Error:",
			err && err.message ? err.message : err
		);
		process.exitCode = 1;
	} finally {
		await mongoose.disconnect();
	}
}

main();
