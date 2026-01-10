const mongoose = require("mongoose");

const reportSchema = new mongoose.Schema(
	{
		// Appendix items to be appended to the generated report (images and PDFs converted to images)
		appendixItems: {
			type: [
				new mongoose.Schema(
					{
						kind: {
							type: String,
							enum: ["image", "pdf"],
							required: true,
						},
						originalName: { type: String },
						originalPath: { type: String }, // absolute or repo-relative path to uploaded image/PDF
						thumbPath: { type: String }, // optional thumbnail path (first page for PDFs)
						pageImages: { type: [String], default: [] }, // per-page image paths for PDFs; empty for single images
						pageCount: { type: Number, default: 0 },
						order: { type: Number, default: 0 },
						uploadedBy: {
							type: mongoose.Schema.Types.ObjectId,
							ref: "User",
						},
						createdAt: { type: Date, default: Date.now },
					},
					{ _id: true }
				),
			],
			default: [],
		},
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
		// Parcel search data (saved for quick access on report edit)
		sbpiIdNo: {
			type: Number,
		},
		parcelDetails: {
			type: mongoose.Schema.Types.Mixed,
		},
		parcelFetchedAt: {
			type: Date,
		},
		parcelSearchParams: {
			type: new mongoose.Schema(
				{
					distCode: { type: Number },
					vilCode: { type: Number },
					qrtrCode: { type: Number },
					sheet: { type: String },
					planNbr: { type: String },
					parcelNbr: { type: String },
				},
				{ _id: false }
			),
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
