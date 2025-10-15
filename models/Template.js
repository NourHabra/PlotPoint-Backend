const mongoose = require("mongoose");

// Content Block Schema
const contentBlockSchema = new mongoose.Schema({
	id: {
		type: String,
		required: true,
	},
	type: {
		type: String,
		enum: ["text", "variable", "kml_variable"],
		required: true,
	},
	content: {
		type: String,
		required: true,
	},
	variableName: {
		type: String,
	},
	variableType: {
		type: String,
		enum: ["string", "number", "date", "currency"],
	},
	kmlField: {
		type: String,
		enum: [
			"municipality",
			"plot_number",
			"plot_area",
			"coordinates",
			"sheet_plan",
			"registration_number",
			"property_type",
			"zone",
			"zone_description",
			"building_coefficient",
			"coverage",
			"floors",
			"height",
			"value_2018",
			"value_2021",
		],
	},
	// New: predefined text templates for variable blocks
	textTemplates: [{ type: String }],
	isRequired: {
		type: Boolean,
		default: false,
	},
	placeholder: {
		type: String,
	},
	validation: {
		minLength: Number,
		maxLength: Number,
		pattern: String,
		min: Number,
		max: Number,
	},
});

// Template Section Schema
const templateSectionSchema = new mongoose.Schema({
	id: {
		type: String,
		required: true,
	},
	title: {
		type: String,
		required: true,
	},
	content: [contentBlockSchema],
	order: {
		type: Number,
		required: true,
	},
});

// Template Schema
const templateSchema = new mongoose.Schema(
	{
		name: {
			type: String,
			required: true,
		},
		description: {
			type: String,
		},
		requiresKml: {
			type: Boolean,
			default: false,
		},
		sections: [templateSectionSchema],
		// Imported DOCX source (if provided via import flow)
		sourceDocxPath: {
			type: String,
		},
		// Cached unfilled PDF preview path for quick stage-1 preview
		previewPdfPath: {
			type: String,
			default: "",
		},
		// Variables metadata for imported templates
		variables: [
			new mongoose.Schema(
				{
					id: { type: String, required: true },
					name: { type: String, required: true },
					type: {
						type: String,
						enum: [
							"text",
							"kml",
							"image",
							"select",
							"date",
							"calculated",
						],
						required: true,
					},
					description: { type: String },
					sourceText: { type: String },
					kmlField: { type: String },
					options: [{ type: String }],
					expression: { type: String },
					// New: predefined text templates for text variables
					textTemplates: [{ type: String }],
					// Image-specific mapping info for replacing media in DOCX
					imageRelId: { type: String }, // rId of the picture relationship in document.xml.rels
					imageTarget: { type: String }, // e.g., media/image3.png
					imageExtent: {
						cx: { type: Number }, // EMUs
						cy: { type: Number }, // EMUs
					},
					// Optional grouping
					groupId: { type: String },
					isRequired: { type: Boolean, default: false },
					tokenized: { type: Boolean, default: false },
				},
				{ _id: false }
			),
		],
		// Optional variable groups for imported templates
		variableGroups: [
			new mongoose.Schema(
				{
					id: { type: String, required: true },
					name: { type: String, required: true },
					description: { type: String },
					order: { type: Number, default: 0 },
				},
				{ _id: false }
			),
		],
		createdAt: {
			type: Date,
			default: Date.now,
		},
		updatedAt: {
			type: Date,
			default: Date.now,
		},
		createdBy: {
			type: String,
			required: true,
		},
		isActive: {
			type: Boolean,
			default: true,
		},
	},
	{
		timestamps: true, // This will automatically handle createdAt and updatedAt
	}
);

// Update the updatedAt field before saving
templateSchema.pre("save", function (next) {
	this.updatedAt = new Date();
	next();
});

module.exports = mongoose.model("Template", templateSchema);
