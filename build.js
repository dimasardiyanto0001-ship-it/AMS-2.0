#!/usr/bin/env node
/**
 * build.js - menyusun halaman Netlify dari file sumber Apps Script.
 *
 *   node build.js           -> target PRODUKSI (build.config.json: "prod")
 *   node build.js --test    -> target SALINAN/testing ("test") + banner merah
 *   AMS_TARGET=test node build.js   -> sama dengan --test (dipakai netlify.toml)
 *
 * Input  : index.html (sumber, masih berisi <?!= include('X'); ?>),
 *          Stylesheet.html, SettingJadwalKerja.html, JavaScript.html, AbsensiOnline.html, FormIzin.html
 * Output : public/index.html, public/absensi-online.html, public/izin.html  (folder public/ yang dipublikasikan Netlify)
 *
 * File sumber TIDAK diubah, jadi tetap bisa ditempel apa adanya ke editor Apps Script.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = __dirname;
const OUT = path.join(ROOT, 'public');
const target = (process.argv.includes('--test') || process.env.AMS_TARGET === 'test') ? 'test' : 'prod';
const URL_RE = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]{20,}\/exec$/;

function fail(msg) { console.error('BUILD GAGAL: ' + msg); process.exit(1); }
function read(name) {
  const p = path.join(ROOT, name);
  if (!fs.existsSync(p)) fail('file tidak ditemukan: ' + name);
  return fs.readFileSync(p, 'utf8');
}

const config = JSON.parse(read('build.config.json'));
const gasUrl = config[target];
if (!gasUrl || !URL_RE.test(gasUrl)) fail('URL "' + target + '" di build.config.json bukan URL web app /exec yang valid');

function inlineIncludes(html, seen) {
  seen = seen || [];
  return html.replace(/<\?!=\s*include\(\s*'([^']+)'\s*\)\s*;?\s*\?>/g, function (_, name) {
    if (seen.indexOf(name) !== -1) fail('include melingkar: ' + name);
    return inlineIncludes(read(name + '.html'), seen.concat(name));
  });
}

function setGasUrl(html, label) {
  const re = /(GAS_WEBAPP_URL_\s*=\s*)'[^']*'/g;
  const n = (html.match(re) || []).length;
  if (n !== 1) fail(label + ': GAS_WEBAPP_URL_ ditemukan ' + n + ' kali (seharusnya 1)');
  return html.replace(re, function (_, pre) { return pre + "'" + gasUrl + "'"; });
}

function check(html, label) {
  if (/<\?/.test(html)) fail(label + ': masih ada tag <? (scriptlet Apps Script) yang belum diproses');
  const re = /<script(?![^>]*\bsrc\b)[^>]*>([\s\S]*?)<\/script>/g;
  let m, i = 0;
  while ((m = re.exec(html))) {
    i++;
    try { new vm.Script(m[1], { filename: label + '#script' + i }); }
    catch (e) { fail(label + ': kesalahan sintaks di blok <script> #' + i + ': ' + e.message); }
  }
  return i;
}

function banner(html) {
  const bar = '<div style="position:fixed;left:0;right:0;bottom:0;z-index:2147483647;background:#B91C1C;color:#fff;' +
    'font:600 12px/1.2 system-ui,sans-serif;text-align:center;padding:4px 8px;pointer-events:none">' +
    'MODE TEST - terhubung ke sheet SALINAN, bukan produksi</div>';
  if (!/<body[^>]*>/i.test(html)) fail('tag <body> tidak ditemukan untuk banner test');
  return html.replace(/<body[^>]*>/i, function (b) { return b + bar; });
}

fs.mkdirSync(OUT, { recursive: true });
const pages = [
  { src: 'index.html', out: 'index.html', inline: true },
  { src: 'AbsensiOnline.html', out: 'absensi-online.html', inline: true },
  { src: 'FormIzin.html', out: 'izin.html', inline: true }
];
pages.forEach(function (p) {
  let html = read(p.src);
  if (p.inline) html = inlineIncludes(html);
  html = setGasUrl(html, p.out);
  if (target === 'test') html = banner(html);
  const scripts = check(html, p.out);
  fs.writeFileSync(path.join(OUT, p.out), html);
  console.log('  ' + p.out.padEnd(22) + (html.length / 1024).toFixed(0).padStart(5) + ' KB, ' + scripts + ' blok script OK');
});
console.log('Target: ' + target.toUpperCase() + ' -> ...' + gasUrl.slice(-24));

// --- PWA: salin manifest, service worker, dan ikon ke folder public ---
{
  const _fs = require('fs');
  const _path = require('path');
  const out = _path.join(__dirname, 'public');
  _fs.mkdirSync(_path.join(out, 'icons'), { recursive: true });
  for (const f of ['manifest.json', 'sw.js']) {
    _fs.copyFileSync(_path.join(__dirname, f), _path.join(out, f));
  }
  for (const f of ['icon-192.png', 'icon-512.png', 'icon-maskable-512.png']) {
    _fs.copyFileSync(_path.join(__dirname, f), _path.join(out, 'icons', f));
  }
}


// --- PWA: assetlinks.json agar bar alamat hilang di aplikasi ---
{
  const _fs = require('fs');
  const _path = require('path');
  const dir = _path.join(__dirname, 'public', '.well-known');
  _fs.mkdirSync(dir, { recursive: true });
  _fs.copyFileSync(
    _path.join(__dirname, 'assetlinks.json'),
    _path.join(dir, 'assetlinks.json')
  );
}
