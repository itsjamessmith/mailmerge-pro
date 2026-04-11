/**
 * MailMerge-Pro Build Script
 * Minifies JS, CSS, and HTML for production deployment.
 * Usage:
 *   node build.js          — Production build (minified output in dist/)
 *   node build.js --lint   — Syntax check only (no output)
 *   node build.js --watch  — Watch mode for development
 */
const fs = require('fs');
const path = require('path');

const DIST = path.join(__dirname, 'dist');
const SRC = __dirname;

const args = process.argv.slice(2);
const isLint = args.includes('--lint');
const isWatch = args.includes('--watch');

async function lint() {
    console.log('🔍 Checking JavaScript syntax...');
    try {
        const esbuild = require('esbuild');
        await esbuild.build({
            entryPoints: [path.join(SRC, 'taskpane.js')],
            bundle: false,
            write: false,
            logLevel: 'error',
            target: 'es2020'
        });
        console.log('✅ taskpane.js — syntax OK');
    } catch (e) {
        console.error('❌ taskpane.js — syntax errors found');
        process.exit(1);
    }

    // Validate manifest.xml is well-formed
    const manifestSrc = fs.readFileSync(path.join(SRC, 'manifest.xml'), 'utf8');
    if (!manifestSrc.includes('<?xml') || !manifestSrc.includes('</OfficeApp>')) {
        console.error('❌ manifest.xml — malformed XML');
        process.exit(1);
    }
    console.log('✅ manifest.xml — well-formed');

    // Validate HTML files parse correctly
    for (const htmlFile of ['taskpane.html', 'index.html', 'function-file.html']) {
        const htmlPath = path.join(SRC, htmlFile);
        if (fs.existsSync(htmlPath)) {
            const content = fs.readFileSync(htmlPath, 'utf8');
            if (!content.includes('<!DOCTYPE html') && !content.includes('<html')) {
                console.error(`❌ ${htmlFile} — missing DOCTYPE or html tag`);
                process.exit(1);
            }
            console.log(`✅ ${htmlFile} — structure OK`);
        }
    }

    console.log('\n✅ All lint checks passed.');
}

async function build() {
    console.log('🏗️  Building MailMerge-Pro...\n');

    // Clean dist
    if (fs.existsSync(DIST)) fs.rmSync(DIST, { recursive: true });
    fs.mkdirSync(DIST, { recursive: true });
    fs.mkdirSync(path.join(DIST, 'assets'), { recursive: true });

    // 1. Minify JS with esbuild
    const esbuild = require('esbuild');
    const jsResult = await esbuild.build({
        entryPoints: [path.join(SRC, 'taskpane.js')],
        outfile: path.join(DIST, 'taskpane.js'),
        minify: true,
        target: 'es2020',
        format: 'iife',
        sourcemap: true,
        logLevel: 'info'
    });
    const jsSrc = fs.statSync(path.join(SRC, 'taskpane.js')).size;
    const jsOut = fs.statSync(path.join(DIST, 'taskpane.js')).size;
    console.log(`  JS: ${(jsSrc / 1024).toFixed(1)} KB → ${(jsOut / 1024).toFixed(1)} KB (${((1 - jsOut / jsSrc) * 100).toFixed(0)}% reduction)`);

    // 2. Minify CSS with clean-css
    const CleanCSS = require('clean-css');
    const cssSrc = fs.readFileSync(path.join(SRC, 'taskpane.css'), 'utf8');
    const cssOut = new CleanCSS({ level: 2 }).minify(cssSrc);
    if (cssOut.errors.length) {
        console.error('❌ CSS errors:', cssOut.errors);
        process.exit(1);
    }
    fs.writeFileSync(path.join(DIST, 'taskpane.css'), cssOut.styles);
    console.log(`  CSS: ${(Buffer.byteLength(cssSrc) / 1024).toFixed(1)} KB → ${(Buffer.byteLength(cssOut.styles) / 1024).toFixed(1)} KB (${((1 - Buffer.byteLength(cssOut.styles) / Buffer.byteLength(cssSrc)) * 100).toFixed(0)}% reduction)`);

    // 3. Minify HTML with html-minifier-terser
    const { minify } = require('html-minifier-terser');
    for (const htmlFile of ['taskpane.html', 'index.html', 'function-file.html']) {
        const htmlPath = path.join(SRC, htmlFile);
        if (fs.existsSync(htmlPath)) {
            const htmlSrc = fs.readFileSync(htmlPath, 'utf8');
            const htmlOut = await minify(htmlSrc, {
                collapseWhitespace: true,
                removeComments: true,
                removeRedundantAttributes: true,
                minifyCSS: true,
                minifyJS: true
            });
            fs.writeFileSync(path.join(DIST, htmlFile), htmlOut);
            console.log(`  ${htmlFile}: ${(Buffer.byteLength(htmlSrc) / 1024).toFixed(1)} KB → ${(Buffer.byteLength(htmlOut) / 1024).toFixed(1)} KB`);
        }
    }

    // 4. Copy static assets
    fs.copyFileSync(path.join(SRC, 'manifest.xml'), path.join(DIST, 'manifest.xml'));
    const assetsDir = path.join(SRC, 'assets');
    if (fs.existsSync(assetsDir)) {
        for (const file of fs.readdirSync(assetsDir)) {
            fs.copyFileSync(path.join(assetsDir, file), path.join(DIST, 'assets', file));
        }
    }
    console.log('  Copied: manifest.xml, assets/');

    // Summary
    let totalSrc = 0, totalDist = 0;
    for (const f of ['taskpane.js', 'taskpane.css', 'taskpane.html']) {
        if (fs.existsSync(path.join(SRC, f))) totalSrc += fs.statSync(path.join(SRC, f)).size;
        if (fs.existsSync(path.join(DIST, f))) totalDist += fs.statSync(path.join(DIST, f)).size;
    }
    console.log(`\n✅ Build complete → dist/`);
    console.log(`   Total: ${(totalSrc / 1024).toFixed(1)} KB → ${(totalDist / 1024).toFixed(1)} KB (${((1 - totalDist / totalSrc) * 100).toFixed(0)}% smaller)`);
}

async function watch() {
    console.log('👀 Watching for changes...\n');
    const files = ['taskpane.js', 'taskpane.css', 'taskpane.html'];
    for (const f of files) {
        fs.watch(path.join(SRC, f), { persistent: true }, async (eventType) => {
            if (eventType === 'change') {
                console.log(`\n📝 ${f} changed — rebuilding...`);
                try { await build(); } catch (e) { console.error(e.message); }
            }
        });
    }
    await build();
}

(async () => {
    try {
        if (isLint) await lint();
        else if (isWatch) await watch();
        else await build();
    } catch (e) {
        console.error('Build failed:', e.message);
        process.exit(1);
    }
})();
