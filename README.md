# Template Backend API

A simple Express.js backend with Mongoose for storing and managing report templates.

## Setup

1. Install dependencies:
```bash
npm install
```

2. Create a `.env` file based on `env.example`:
```bash
cp env.example .env
```

3. Update the `.env` file with your MongoDB connection string.

4. Start the server:
```bash
# Development mode
npm run dev

# Production mode
npm start
```

## API Endpoints

- `GET /api/templates` - Get all active templates
- `GET /api/templates/:id` - Get template by ID
- `POST /api/templates` - Create new template
- `PUT /api/templates/:id` - Update template
- `DELETE /api/templates/:id` - Soft delete template

## Template Schema

The template schema supports:
- Template metadata (name, description, category)
- Sections with ordered content blocks
- Three types of content blocks: text, variable, and kml_variable
- Validation rules for variables
- KML field integration
- Soft deletion with isActive flag

## Example Template Structure

```json
{
  "name": "Property Valuation Report",
  "description": "Standard template for property valuation",
  "category": "Field Report",
  "createdBy": "user123",
  "sections": [
    {
      "id": "section-1",
      "title": "Property Information",
      "order": 0,
      "content": [
        {
          "id": "block-1",
          "type": "text",
          "content": "The property located in "
        },
        {
          "id": "block-2",
          "type": "variable",
          "content": "{{property_location}}",
          "variableName": "property_location",
          "variableType": "string"
        }
      ]
    }
  ]
}
``` 