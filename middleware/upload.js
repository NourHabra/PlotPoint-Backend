const multer = require('multer');
const path = require('path');
const fs = require('fs');

// Allowed file types with their MIME types
const ALLOWED_IMAGE_TYPES = {
  'image/jpeg': ['.jpg', '.jpeg'],
  'image/png': ['.png'],
  'image/gif': ['.gif'],
  'image/webp': ['.webp'],
};

const ALLOWED_DOCUMENT_TYPES = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
  'application/msword': ['.doc'],
};

const ALLOWED_AVATAR_TYPES = {
  'image/jpeg': ['.jpg', '.jpeg'],
  'image/png': ['.png'],
};

// File size limits (in bytes)
const MAX_IMAGE_SIZE = 100 * 1024 * 1024; // 100MB
const MAX_DOCUMENT_SIZE = 100 * 1024 * 1024; // 100MB
const MAX_AVATAR_SIZE = 100 * 1024 * 1024; // 100MB

// Validate file extension matches MIME type
const validateFileType = (file, allowedTypes) => {
  const mimetype = file.mimetype.toLowerCase();
  const extname = path.extname(file.originalname).toLowerCase();
  
  if (!allowedTypes[mimetype]) {
    return false;
  }
  
  if (!allowedTypes[mimetype].includes(extname)) {
    return false;
  }
  
  return true;
};

// Sanitize filename to prevent path traversal
const sanitizeFilename = (filename) => {
  // Remove path separators and special characters
  return filename
    .replace(/[\/\\]/g, '')
    .replace(/[^a-zA-Z0-9._-]/g, '_')
    .substring(0, 100); // Limit filename length
};

// Image upload configuration
const createImageUpload = (uploadDir) => {
  // Ensure upload directory exists
  if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
  }

  const storage = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
      const sanitized = sanitizeFilename(file.originalname);
      const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
      const ext = path.extname(sanitized);
      const basename = path.basename(sanitized, ext);
      cb(null, basename + '-' + uniqueSuffix + ext);
    }
  });

  return multer({
    storage,
    limits: {
      fileSize: MAX_IMAGE_SIZE,
      files: 1,
    },
    fileFilter: (req, file, cb) => {
      if (validateFileType(file, ALLOWED_IMAGE_TYPES)) {
        cb(null, true);
      } else {
        cb(new Error('Invalid file type. Only JPEG, PNG, GIF, and WebP images are allowed.'));
      }
    }
  });
};

// Document upload configuration
const createDocumentUpload = (uploadDir) => {
  // Ensure upload directory exists
  if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
  }

  const storage = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
      const sanitized = sanitizeFilename(file.originalname);
      const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
      const ext = path.extname(sanitized);
      const basename = path.basename(sanitized, ext);
      cb(null, basename + '-' + uniqueSuffix + ext);
    }
  });

  return multer({
    storage,
    limits: {
      fileSize: MAX_DOCUMENT_SIZE,
      files: 1,
    },
    fileFilter: (req, file, cb) => {
      if (validateFileType(file, ALLOWED_DOCUMENT_TYPES)) {
        cb(null, true);
      } else {
        cb(new Error('Invalid file type. Only DOCX and DOC files are allowed.'));
      }
    }
  });
};

// Avatar upload configuration
const createAvatarUpload = (uploadDir) => {
  // Ensure upload directory exists
  if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
  }

  const storage = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
      const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
      const ext = path.extname(file.originalname).toLowerCase();
      cb(null, 'avatar-' + uniqueSuffix + ext);
    }
  });

  return multer({
    storage,
    limits: {
      fileSize: MAX_AVATAR_SIZE,
      files: 1,
    },
    fileFilter: (req, file, cb) => {
      if (validateFileType(file, ALLOWED_AVATAR_TYPES)) {
        cb(null, true);
      } else {
        cb(new Error('Invalid file type. Only JPEG and PNG images are allowed for avatars.'));
      }
    }
  });
};

// Appendix upload configuration (multiple files)
const createAppendixUpload = (uploadDir) => {
  // Ensure upload directory exists
  if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
  }

  const storage = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
      const sanitized = sanitizeFilename(file.originalname);
      const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
      const ext = path.extname(sanitized);
      const basename = path.basename(sanitized, ext);
      cb(null, basename + '-' + uniqueSuffix + ext);
    }
  });

  return multer({
    storage,
    limits: {
      fileSize: MAX_IMAGE_SIZE,
      files: 20, // Allow multiple files for appendix
    },
    fileFilter: (req, file, cb) => {
      if (validateFileType(file, ALLOWED_IMAGE_TYPES)) {
        cb(null, true);
      } else {
        cb(new Error('Invalid file type. Only JPEG, PNG, GIF, and WebP images are allowed.'));
      }
    }
  });
};

module.exports = {
  createImageUpload,
  createDocumentUpload,
  createAvatarUpload,
  createAppendixUpload,
  sanitizeFilename,
};

