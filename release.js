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
    rl.close(); process.exit(1);
  }

  rl.question('Describe what changed (patch notes): ', (desc) => {
    desc = desc.trim() || 'Minor improvements and bug fixes';
    rl.close();

    console.log(`\n📦 Releasing v${version} — "${desc}"\n`);

    try {
      pkg.version = version;
      fs.writeFileSync('package.json', JSON.stringify(pkg, null, 2) + '\n');
      console.log(`✓ Updated package.json to v${version}`);

      execSync('git add .', { stdio: 'inherit' });
      execSync(`git commit -m "v${version} - ${desc}"`, { stdio: 'inherit' });
      console.log('✓ git commit');

      execSync('git push', { stdio: 'inherit' });
      console.log('✓ git push');

      console.log('\n🔨 Building and publishing to GitHub...\n');
      execSync('npm run electron:build', { stdio: 'inherit' });

      console.log(`\n✅ v${version} released successfully!`);

      // Clean up old installers
      cleanupDist(version);

      // Upload patch notes to GitHub release body
      uploadPatchNotes(version, desc);

    } catch (err) {
      console.error('\n❌ Release failed:', err.message);
      process.exit(1);
    }
  });
});

function uploadPatchNotes(version, notes) {
  const token = process.env.GH_TOKEN;
  if (!token) { console.log('\n⚠  No GH_TOKEN found — patch notes not uploaded'); return; }

  const tag = `v${version}`;
  console.log('\n📝 Uploading patch notes to GitHub release...');

  // Get release by tag
  githubRequest('GET', `/repos/vslnnd/Splitify/releases/tags/${tag}`, token, null, (err, release) => {
    if (err || !release || !release.id) {
      console.log('⚠  Could not find GitHub release to update — patch notes skipped');
      return;
    }
    // Update release body with patch notes
    githubRequest('PATCH', `/repos/vslnnd/Splitify/releases/${release.id}`, token, { body: notes }, (err2) => {
      if (err2) console.log('⚠  Failed to update release notes:', err2.message);
      else console.log('✓  Patch notes uploaded to GitHub release');
    });
  });
}

function githubRequest(method, endpoint, token, body, cb) {
  const data = body ? JSON.stringify(body) : null;
  const options = {
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
  };
  const req = https.request(options, (res) => {
    let raw = '';
    res.on('data', chunk => raw += chunk);
    res.on('end', () => {
      try { cb(null, JSON.parse(raw)); }
      catch(e) { cb(null, {}); }
    });
  });
  req.on('error', (e) => cb(e));
  if (data) req.write(data);
  req.end();
}

function cleanupDist(newVersion) {
  const distDir = path.join(__dirname, 'dist');
  if (!fs.existsSync(distDir)) return;
  const newExe = `Splitify Setup ${newVersion}.exe`;
  const newBm  = `Splitify Setup ${newVersion}.exe.blockmap`;
  if (!fs.existsSync(path.join(distDir, newExe))) {
    console.log('\n⚠  New installer not found in dist/ — skipping cleanup.'); return;
  }
  let deleted = 0;
  fs.readdirSync(distDir).forEach(file => {
    const isOldExe = file.endsWith('.exe')          && file.startsWith('Splitify Setup') && file !== newExe;
    const isOldBm  = file.endsWith('.exe.blockmap') && file.startsWith('Splitify Setup') && file !== newBm;
    if (isOldExe || isOldBm) {
      try { fs.unlinkSync(path.join(distDir, file)); console.log(`🗑  Removed: ${file}`); deleted++; }
      catch(e) { console.warn(`⚠  Could not remove ${file}: ${e.message}`); }
    }
  });
  if (deleted > 0) console.log(`\n✓ Cleaned up ${deleted} old installer file${deleted !== 1 ? 's' : ''} from dist/`);
  else console.log('\n✓ dist/ already clean');
}
