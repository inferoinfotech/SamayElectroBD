const fs = require('fs');
const path = require('path');

const PUBLIC_IMAGES_DIR = path.join(__dirname, '../../public/images');

const getMimeType = (filePath) => {
    const ext = path.extname(filePath).toLowerCase();
    const map = {
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.svg': 'image/svg+xml',
        '.webp': 'image/webp',
    };
    return map[ext] || 'application/octet-stream';
};

/** Map an img src URL to a local public/images file when possible. */
const resolveLocalImagePath = (src) => {
    if (!src || src.startsWith('cid:') || src.startsWith('data:')) {
        return null;
    }

    const fileNameMatch = src.match(/\/public\/images\/([^/?#]+)/i);
    if (!fileNameMatch) {
        return null;
    }

    const fileName = decodeURIComponent(fileNameMatch[1]);
    const filePath = path.join(PUBLIC_IMAGES_DIR, fileName);

    if (fs.existsSync(filePath)) {
        return filePath;
    }

    // Template may reference logo.png while server has logo.jpg
    const base = path.basename(fileName, path.extname(fileName));
    for (const ext of ['.jpg', '.jpeg', '.png', '.webp', '.svg', '.gif']) {
        const altPath = path.join(PUBLIC_IMAGES_DIR, `${base}${ext}`);
        if (fs.existsSync(altPath)) {
            return altPath;
        }
    }

    return null;
};

/** Should this src be embedded inline (localhost / our backend), not fetched by recipient mail client? */
const shouldEmbedInline = (src) => {
    if (!src || src.startsWith('cid:') || src.startsWith('data:')) {
        return false;
    }

    if (/\/public\/images\//i.test(src)) {
        if (/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?\//i.test(src)) {
            return true;
        }

        const backendUrl = (process.env.BACKEND_URL || '').replace(/\/$/, '');
        if (backendUrl && src.startsWith(backendUrl)) {
            return true;
        }

        // Relative or same-server paths used in saved templates
        if (src.startsWith('/public/images/')) {
            return true;
        }
    }

    return false;
};

/**
 * Replace local/backend image URLs in HTML with CID references and attach files.
 * Used for general emails so Gmail/Outlook can show logos without hitting localhost.
 */
const embedInlineImagesForEmail = (html, existingAttachments = []) => {
    if (!html) {
        return { html: html || '', attachments: existingAttachments };
    }

    const attachments = [...existingAttachments];
    let cidCounter = 0;

    const processedHtml = html.replace(
        /<img\b([^>]*?)\bsrc=["']([^"']+)["']([^>]*)>/gi,
        (fullTag, before, src, after) => {
            if (!shouldEmbedInline(src)) {
                return fullTag;
            }

            const localPath = resolveLocalImagePath(src);
            if (!localPath) {
                return fullTag;
            }

            cidCounter += 1;
            const cid = `samay-inline-${cidCounter}@samayelectro`;

            attachments.push({
                filename: path.basename(localPath),
                path: localPath,
                cid,
                contentType: getMimeType(localPath),
                contentDisposition: 'inline',
            });

            return `<img${before}src="cid:${cid}"${after}>`;
        }
    );

    return { html: processedHtml, attachments };
};

module.exports = {
    embedInlineImagesForEmail,
    resolveLocalImagePath,
};
