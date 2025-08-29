const fetch = require("node-fetch");

const API_BASE_URL = "http://localhost:5000/api";

async function testTemplateAPI() {
	console.log("Testing Template API...\n");

	// Test data
	const testTemplate = {
		name: "Test Property Report",
		description: "A test template for property reports",
		createdBy: "test-user",
		sections: [
			{
				id: "section-1",
				title: "Property Information",
				order: 0,
				content: [
					{
						id: "block-1",
						type: "text",
						content: "The property located in ",
					},
					{
						id: "block-2",
						type: "variable",
						content: "{{property_location}}",
						variableName: "property_location",
						variableType: "string",
					},
					{
						id: "block-3",
						type: "text",
						content: " has an area of ",
					},
					{
						id: "block-4",
						type: "kml_variable",
						content: "{{plot_area}}",
						kmlField: "plot_area",
						variableType: "number",
					},
					{
						id: "block-5",
						type: "text",
						content: " square meters.",
					},
				],
			},
		],
	};

	try {
		// Test 1: Create template
		console.log("1. Creating template...");
		const createResponse = await fetch(`${API_BASE_URL}/templates`, {
			method: "POST",
			headers: {
				"Content-Type": "application/json",
			},
			body: JSON.stringify(testTemplate),
		});

		if (!createResponse.ok) {
			throw new Error(
				`Create failed: ${createResponse.status} ${createResponse.statusText}`
			);
		}

		const createdTemplate = await createResponse.json();
		console.log("✅ Template created successfully");
		console.log("   ID:", createdTemplate._id);
		console.log("   Name:", createdTemplate.name);

		// Test 2: Get all templates
		console.log("\n2. Getting all templates...");
		const getAllResponse = await fetch(`${API_BASE_URL}/templates`);

		if (!getAllResponse.ok) {
			throw new Error(
				`Get all failed: ${getAllResponse.status} ${getAllResponse.statusText}`
			);
		}

		const allTemplates = await getAllResponse.json();
		console.log("✅ Retrieved templates successfully");
		console.log("   Count:", allTemplates.length);

		// Test 3: Get specific template
		console.log("\n3. Getting specific template...");
		const getOneResponse = await fetch(
			`${API_BASE_URL}/templates/${createdTemplate._id}`
		);

		if (!getOneResponse.ok) {
			throw new Error(
				`Get one failed: ${getOneResponse.status} ${getOneResponse.statusText}`
			);
		}

		const retrievedTemplate = await getOneResponse.json();
		console.log("✅ Retrieved specific template successfully");
		console.log("   Name:", retrievedTemplate.name);
		console.log("   Sections:", retrievedTemplate.sections.length);

		// Test 4: Update template
		console.log("\n4. Updating template...");
		const updateData = {
			...testTemplate,
			name: "Updated Test Property Report",
			description: "Updated description",
		};

		const updateResponse = await fetch(
			`${API_BASE_URL}/templates/${createdTemplate._id}`,
			{
				method: "PUT",
				headers: {
					"Content-Type": "application/json",
				},
				body: JSON.stringify(updateData),
			}
		);

		if (!updateResponse.ok) {
			throw new Error(
				`Update failed: ${updateResponse.status} ${updateResponse.statusText}`
			);
		}

		const updatedTemplate = await updateResponse.json();
		console.log("✅ Template updated successfully");
		console.log("   New name:", updatedTemplate.name);

		// Test 5: Delete template (soft delete)
		console.log("\n5. Deleting template...");
		const deleteResponse = await fetch(
			`${API_BASE_URL}/templates/${createdTemplate._id}`,
			{
				method: "DELETE",
			}
		);

		if (!deleteResponse.ok) {
			throw new Error(
				`Delete failed: ${deleteResponse.status} ${deleteResponse.statusText}`
			);
		}

		console.log("✅ Template deleted successfully");

		// Test 6: Verify template is not in active list
		console.log("\n6. Verifying template is inactive...");
		const finalGetResponse = await fetch(`${API_BASE_URL}/templates`);
		const finalTemplates = await finalGetResponse.json();

		const deletedTemplate = finalTemplates.find(
			(t) => t._id === createdTemplate._id
		);
		if (deletedTemplate) {
			console.log("❌ Template still appears in active list");
		} else {
			console.log("✅ Template correctly removed from active list");
		}

		// Test 7: Get all templates including inactive
		console.log("\n7. Getting all templates including inactive...");
		const getAllInactiveResponse = await fetch(
			`${API_BASE_URL}/templates?includeInactive=true`
		);

		if (!getAllInactiveResponse.ok) {
			throw new Error(
				`Get all inactive failed: ${getAllInactiveResponse.status} ${getAllInactiveResponse.statusText}`
			);
		}

		const allInactiveTemplates = await getAllInactiveResponse.json();
		const inactiveTemplate = allInactiveTemplates.find(
			(t) => t._id === createdTemplate._id
		);
		console.log("✅ Retrieved all templates including inactive");
		console.log("   Total count:", allInactiveTemplates.length);
		console.log("   Inactive template found:", !!inactiveTemplate);

		// Test 8: Reactivate template
		console.log("\n8. Reactivating template...");
		const reactivateResponse = await fetch(
			`${API_BASE_URL}/templates/${createdTemplate._id}/reactivate`,
			{
				method: "PATCH",
			}
		);

		if (!reactivateResponse.ok) {
			throw new Error(
				`Reactivate failed: ${reactivateResponse.status} ${reactivateResponse.statusText}`
			);
		}

		const reactivatedTemplate = await reactivateResponse.json();
		console.log("✅ Template reactivated successfully");
		console.log(
			"   Template active:",
			reactivatedTemplate.template.isActive
		);

		// Test 9: Verify template is back in active list
		console.log("\n9. Verifying template is back in active list...");
		const finalActiveResponse = await fetch(`${API_BASE_URL}/templates`);
		const finalActiveTemplates = await finalActiveResponse.json();

		const reactivatedTemplateInList = finalActiveTemplates.find(
			(t) => t._id === createdTemplate._id
		);
		if (reactivatedTemplateInList) {
			console.log("✅ Template correctly restored to active list");
		} else {
			console.log(
				"❌ Template not found in active list after reactivation"
			);
		}

		console.log("\n🎉 All tests passed!");
	} catch (error) {
		console.error("❌ Test failed:", error.message);
		process.exit(1);
	}
}

// Run the test
testTemplateAPI();
