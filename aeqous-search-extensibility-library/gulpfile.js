'use strict';

const gulp = require('gulp');
const log = require('fancy-log');
const fs = require('fs');

const readJson = (path, cb) => {
    fs.readFile(require.resolve(path), (err, data) => {
        if (err)
        log.error(err)
        else
        cb(null, JSON.parse(data))
    });
}

const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(gulp);

gulp.task('update-version', async () => {

    const libraryManifests = [
      './src/libraries/aeqousSearchLibrary/AeqousSearchLibrary.manifest.json'
    ];

    const semver = require('semver')
    const versionArgIdx = process.argv.indexOf('--value');
    const newVersionNumber = semver.valid(process.argv[versionArgIdx+1]);

    if (versionArgIdx !== -1 && newVersionNumber) {

        // Update version in the package-solution
        const pkgSolutionFilePath = './config/package-solution.json';

        readJson(pkgSolutionFilePath, (err, pkgSolution) => {
          log.info('Old package-solution.json version:\t' + pkgSolution.solution.version);
          const pkgVersion = `${semver.major(newVersionNumber)}.${semver.minor(newVersionNumber)}.${semver.patch(newVersionNumber)}.0`;
          pkgSolution.solution.version = pkgVersion
          log.info('New package-solution.json version:\t' + pkgVersion);
          fs.writeFile(pkgSolutionFilePath, JSON.stringify(pkgSolution, null, 4), (error) => {});
        });

        // Updated version in library manifests
        libraryManifests.forEach((manifestFile) => {
          readJson(manifestFile, (err, manifest) => {

            log.info('Old manifestFile version:\t' + manifest.version);
            const libVersion = `${semver.major(newVersionNumber)}.${semver.minor(newVersionNumber)}.${semver.patch(newVersionNumber)}`;
            manifest.version = libVersion;
            log.info('New manifestFile version:\t' + libVersion);
            fs.writeFile(manifestFile, JSON.stringify(manifest, null, 4), (error) => {});
          });
        });
    } else {
        log.error(`The provided version ${process.argv[versionArgIdx+1]} is not a valid SemVer version`);
    }
  });
