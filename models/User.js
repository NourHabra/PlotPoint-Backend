const mongoose = require("mongoose");

const userSchema = new mongoose.Schema(
	{
		name: { type: String, required: true },
		email: { type: String, required: true, unique: true, index: true },
		passwordHash: { type: String, required: true },
		role: { type: String, enum: ["Admin", "User"], default: "User" },
		avatarPath: { type: String, default: "" },
	},
	{ timestamps: true }
);

module.exports = mongoose.model("User", userSchema);
