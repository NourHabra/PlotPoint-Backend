const mongoose = require("mongoose");

const changeLogSchema = new mongoose.Schema(
	{
		title: { type: String, required: true },
		description: { type: String, required: true },
		date: { type: Date, required: true },
		disabled: { type: Boolean, default: false },
		createdBy: { type: String, default: "system" },
	},
	{ timestamps: true }
);

module.exports = mongoose.model("ChangeLog", changeLogSchema);
