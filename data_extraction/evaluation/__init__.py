"""
Evaluation tools.

``column_evaluator`` evaluates the AI Review *column analysis* output
workbooks against the section content that generated them. It is the
adaptation of the alr evaluation modules kept in this folder as reference:

* ``data_evaluator.py``   — source of the substring/grounding (Is_Subset)
  check and the true/false count bookkeeping (imports ``alr.*``; not
  runnable in this repo as-is).
* ``metric_evaluator.py`` — source of the guarded lexical/distance metric
  batch pattern (also ``alr.*``-bound).
* ``Lexical_Overlap_Metrics.py`` / ``Distance_w_Structural _Alignment.py``
  — the metric implementations (Jaccard/ROUGE/BLEU and Levenshtein/WER),
  used directly by ``column_evaluator`` when their libraries are installed.
"""
