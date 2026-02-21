const { execSync } = require('child_process');
const readline = require('readline');
const fs = require('fs');
const path = require('path');

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
    } catch (err) {
      console.error('\n❌ Release failed:', err.message);
      process.exit(1);
    }
  });
});
