# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

This project is still in the initial development phase, and as such there are numerous
bugs, instabilities and missing features.  See the
[project issues page](https://github.com/bskinn/excel-mregress/issues) for planned
feature enhancements, known bugs, etc.

## [Unreleased]

## [0.1.0] - 2017-08-03

Existing development version released as open source.

### Current Features

 * Multiple linear regressions can be constructed, and various resulting statistics
   are reported
 * Source data is copied before manipulation, so there should be no risk
   of corruption
 * Predictors can be freely masked in/out of the regression analysis, with the resulting
   analysis updated automatically
 * Models can be automatically down-selected based on the Akaike Information Criterion (AIC)
   and its corrected version for small sample sizes (AICc)
 * Various per-datapoint quantities can be plotted against one another (specific predictor values,
   response value, residual, etc.)

### Known Major Bugs

 * Unhandled exception during regression construction if redundant or linearly dependent
   predictors are present
 * Various exceptions occur if data/worksheets/workbooks are not found in expected locations
 * Unhandled exception is raised during regression construction for *very* small models, where
   number of datapoints is insufficient for calculation of the AICc

