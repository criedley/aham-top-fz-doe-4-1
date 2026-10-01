AHAM energy test automation on refrigerators/freezers — it imports raw CSV data cart output, detects events (defrost/unit-on cycles), performs stability checks on temperature gradients, computes energy metrics (EP1/EP2/ET per the AHAM 8.7.2.1.2 formula), and generates a formatted report with charts.*

pipeline: import data → detect cycles → check stability → calculate energy → generate report

*project still in development
