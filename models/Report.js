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
		// Per-report checklist progress (user-defined per template)
		checklistProgress: {
			type: [
				new mongoose.Schema(
					{
						id: { type: String, required: true },
						checked: { type: Boolean, default: false },
					},
					{ _id: false }
				),
			],
			default: [],
		},
		checklistStatus: {
			type: String,
			enum: ["empty", "partial", "complete"],
			default: "empty",
		},
	},
	{ timestamps: true }
);

module.exports = mongoose.model("Report", reportSchema);
