"""Record the two retired legacy-UI checks from the Phase 2 suite.

The unified application deliberately does not vendor the old standalone Phase 2
GUI.  Keeping these test names as explicit skips makes that exception visible in
the regression count without importing the repository-root unified GUI under a
misleading legacy contract.
"""

import pytest


pytestmark = pytest.mark.skip(
    reason="the standalone Phase 2 GUI was replaced by the unified application"
)


def test_discover_batch_docx_excludes_prior_phase2_outputs():
    pass


def test_summarize_batch_results_distinguishes_outcomes():
    pass
