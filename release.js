const { execSync } = require('child_process');
const readline = require('readline');
const fs = require('fs');
const path = require('path');
const https = require('https');

const pkg = JSON.parse(fs.readFileSync('package.json', 'utf8'));

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

console.log('\n🚀 Splitify Release Tool');
console.log('─────────────────────────');
console.log(`Current version: ${pkg.version}`);

rl.question('New version (e.g. 1.1.0): ', (version) => {
  version = version.trim();

  if (!version || !/^\d+\.\d+\.\d+$/.test(version)) {
    console.error('❌ Invalid version format. Use x.y.z (e.g. 1.1.0)');
    rl.close();
    process.exit(1);
  }

  rl.question(`Describe what changed (for commit message): `, (desc) => {
    desc = desc.trim() || 'update';

    rl.question(`Patch notes (shown to users in the app): `, (notes) => {
      notes = notes.trim() || desc;
      rl.close();

      console.log(`\n📦 Releasing v${version} — "${desc}"\n`);

      try {
        // 1. Update version in package.json
        pkg.version = version;
        fs.writeFileSync('package.json', JSON.stringify(pkg, null, 2) + '\n');
        console.log(`✓ Updated package.json to v${version}`);

        // 2. Git add all
        execSync('git add .', { stdio: 'inherit' });
        console.log('✓ git add .');

        // 3. Git commit
        execSync(`git commit -m "v${version} - ${desc}"`, { stdio: 'inherit' });
        console.log(`✓ git commit`);

        // 4. Git push
        execSync('git push', { stdio: 'inherit' });
        console.log('✓ git push');

        // 5. Build + publish
        console.log('\n🔨 Building and publishing to GitHub...\n');
        execSync('npm run electron:build', { stdio: 'inherit' });

        console.log(`\n✅ v${version} released successfully!`);

        // 6. Clean up old installers from dist/ — only after new one confirmed present
        cleanupDist(version);

        // 7. Upload patch notes to GitHub release body
        uploadPatchNotes(version, notes);

      } catch (err) {
        console.error('\n❌ Release failed:', err.message);
        process.exit(1);
      }
    });
  });
});

function cleanupDist(newVersion) {
  const distDir = path.join(__dirname, 'dist');
  if (!fs.existsSync(distDir)) return;

  const newExe      = `Splitify Setup ${newVersion}.exe`;
  const newBlockmap = `Splitify Setup ${newVersion}.exe.blockmap`;

  // Safety check: only clean if the new installer actually exists
  if (!fs.existsSync(path.join(distDir, newExe))) {
    console.log('\n⚠  New installer not found in dist/ — skipping cleanup.');
    return;
  }

  let deleted = 0;
  fs.readdirSync(distDir).forEach(file => {
    const isOldExe      = file.endsWith('.exe')          && file.startsWith('Splitify Setup') && file !== newExe;
    const isOldBlockmap = file.endsWith('.exe.blockmap') && file.startsWith('Splitify Setup') && file !== newBlockmap;

    if (isOldExe || isOldBlockmap) {
      try {
        fs.unlinkSync(path.join(distDir, file));
        console.log(`🗑  Removed: ${file}`);
        deleted++;
      } catch (e) {
        console.warn(`⚠  Could not remove ${file}: ${e.message}`);
      }
    }
  });

  if (deleted > 0) {
    console.log(`\n✓ Cleaned up ${deleted} old installer file${deleted !== 1 ? 's' : ''} from dist/`);
  } else {
    console.log('\n✓ dist/ already clean — no old installers to remove');
  }
}

function uploadPatchNotes(version, notes) {
  const token = process.env.GH_TOKEN;
  if (!token) { console.log('\n⚠  No GH_TOKEN — patch notes not uploaded'); return; }

  console.log('\n📝 Uploading patch notes to GitHub release...');
  githubRequest('GET', `/repos/vslnnd/Splitify/releases/tags/v${version}`, token, null, (err, release) => {
    if (err || !release || !release.id) {
      console.log('⚠  Could not find GitHub release — patch notes skipped'); return;
    }
    githubRequest('PATCH', `/repos/vslnnd/Splitify/releases/${release.id}`, token, { body: notes }, (err2) => {
      if (err2) console.log('⚠  Failed to upload patch notes:', err2.message);
      else console.log('✓  Patch notes uploaded');
    });
  });
}

function githubRequest(method, endpoint, token, body, cb) {
  const data = body ? JSON.stringify(body) : null;
  const req = https.request({
    hostname: 'api.github.com',
    path: endpoint,
    method,
    headers: {
      'Authorization': `token ${token}`,
      'User-Agent': 'Splitify-Release-Tool',
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json',
      ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {})
    }
  }, (res) => {
    let raw = '';
    res.on('data', chunk => raw += chunk);
    res.on('end', () => { try { cb(null, JSON.parse(raw)); } catch(e) { cb(null, {}); } });
  });
  req.on('error', cb);
  if (data) req.write(data);
  req.end();
}
