const mongoose = require("mongoose");

const reportSchema = new mongoose.Schema(
	{
		templateId: {
			type: mongoose.Schema.Types.ObjectId,
			ref: "Template",
			required: true,
		},
		templateName: {
			type: String,
		},
		name: {
			type: String,
		},
		title: {
			type: String,
		},
		status: {
			type: String,
			default: "Draft",
		},
		values: {
			type: mongoose.Schema.Types.Mixed,
			default: {},
		},
		// Optional KML data extracted and stored for report
		kmlData: {
			type: mongoose.Schema.Types.Mixed,
			default: {},
		},
		createdBy: {
			type: String,
			default: "system",
		},
		lastGeneratedAt: {
			type: Date,
		},
		isArchived: {
			type: Boolean,
			default: false,
		},
	},
	{ timestamps: true }
);

module.exports = mongoose.model("Report", reportSchema);
