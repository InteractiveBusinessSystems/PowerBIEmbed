'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// ********* ADDED *******
// disable tslint
build.tslintCmd.enabled = false;
// ********* ADDED *******

build.initialize(require('gulp'));
