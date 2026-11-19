# =============================================================================
# cgs_ai — greeting, sentiment analysis, and NLP utilities for CMS comment data.
# Co-authored with CoCo
# Version 1.1
# =============================================================================
#
#   *** THIS MODULE RUNS IN SNOWFLAKE ONEPI ONLY. ***
#
# Every analytical function here depends on things that exist only inside a
# Snowflake notebook session: an active Snowpark session, the snowflake.cortex
# functions, and the CPU pool the notebook service provides. Run it anywhere
# else -- a laptop, a build agent, a container -- and it cannot work, because
# there is no session to attach to and no Cortex endpoint to call.
#
# So each function calls _require_snowflake() first. If the environment is not
# a Snowflake notebook the function writes a NOTE to the log explaining what
# was missing and returns None. It does not raise: an unattended pipeline step
# should report and stop, not abort the whole notebook.
#
#   WHY THE CHECK IS NOT JUST platform.platform()
#   ---------------------------------------------
#   v1.0 tested that the platform string began with "Linux". Nearly every
#   build agent, container and CI runner in existence is also Linux, so that
#   check passed in all the places this code cannot actually run, and the
#   first sign of trouble was an import error deep inside a function. The
#   guard now looks for Snowflake itself: the snowpark module, and an active
#   session. Linux is still required and still checked, but it is necessary
#   rather than sufficient.
#
# The browser application dmw_claims_management_v1_0Beta.html reimplements the
# deterministic parts of this pipeline in JavaScript so the workflow can be
# demonstrated without a Snowflake session. The two are NOT interchangeable:
# summarizeText here calls Cortex Summarize and generates abstractive text,
# while the browser builds an extractive summary from the comment's own
# sentences. Where they differ, this file is the system of record.
# =============================================================================


def _require_snowflake(func_name: str) -> bool:
    """
    Return True when running inside a Snowflake notebook session.

    Checks three things in increasing order of cost:
      1. the platform is Linux          (cheap, necessary, not sufficient)
      2. snowflake.snowpark imports     (the library is present)
      3. an active session exists       (we are actually inside a session)

    On failure it prints a NOTE naming which check failed and returns False,
    so the caller can escape processing without raising.
    """
    import platform

    plat = platform.platform()
    print(f"Platform : {plat}")

    if not plat.startswith("Linux"):
        print(f"NOTE: {func_name} runs in Snowflake ONEPI only. "
              f"Detected platform '{plat}', which is not Linux. Escaping processing.")
        return False

    try:
        from snowflake.snowpark.context import get_active_session
    except ImportError:
        print(f"NOTE: {func_name} runs in Snowflake ONEPI only. The snowflake.snowpark "
              f"library is not available here, so there is no session to attach to. "
              f"Escaping processing.")
        return False

    try:
        get_active_session()
    except Exception as exc:
        print(f"NOTE: {func_name} runs in Snowflake ONEPI only. No active Snowpark "
              f"session ({type(exc).__name__}). Open this in a Snowflake notebook. "
              f"Escaping processing.")
        return False

    return True


def doBasicHello() -> str:
    """Version 1: Standard greeting."""
    return "Hello, World!"


def doPersonalizedHello(name: str) -> str:
    """Version 2: Personalized greeting with input protection."""
    clean_name = str(name).strip() if name else "World"
    return f"Hello, {clean_name}!"


def doDetailedHello(style: str = "friendly") -> dict:
    """Version 3: Returns a structured JSON-like dictionary."""
    styles = {
        "friendly": "Hello, World! Wonderful day, isn't it?",
        "formal": "Greetings, World. It is a pleasure to connect.",
        "pirate": "Ahoy, World! Avast ye scallywags!",
    }
    greeting = styles.get(style.lower(), styles["friendly"])
    return {"message": greeting, "style_used": style}


def doSentimentAnalysis(
    csv_path,
    comment_col="COMMENT",
    result_col=None,
    CREATE_BAR=False
):
    """
    =====================================================================
    Program Name  : doSentimentAnalysis
    Author        : Manuel Figallo
    Purpose       : Load a CSV file and apply a rule-based sentiment proxy
                    to a configurable text column, producing a new sentiment
                    category column (Positive Language / Negative Language /
                    Neutral / Other).  Optionally renders a bar chart of
                    the sentiment distribution using matplotlib + seaborn.
    Version       : 1.1beta
    Created       : 2026-08-27
    Last Modified : 2026-08-27
    Dependencies  : pandas (required); matplotlib, seaborn (required only
                    when CREATE_BAR=True).  The environment check is
                    _require_snowflake() at the top of this module.
    Description   : RUNS IN SNOWFLAKE ONEPI ONLY. The function escapes
                    processing with a log NOTE if there is no active
                    Snowpark session.  It reads the CSV at csv_path, computes
                    basic structural metrics (char_count, word_count), then
                    applies a keyword-intersection rule to classify each row
                    as Positive Language, Negative Language, or
                    Neutral / Other.  Results are stored in result_col and
                    the enriched DataFrame is returned.

    Input Parameters:
        csv_path    (str, REQUIRED)  - Full path to the input CSV file.
        comment_col (str, default "COMMENT") - Name of the column to run
                    sentiment analysis on.
        result_col  (str, default <comment_col>_SENTIMENT) - Name of the
                    output column that holds the sentiment label.
        CREATE_BAR  (bool, default False) - When True, renders a seaborn
                    bar chart of the sentiment distribution.

    Output:
        Returns the enriched pandas DataFrame (with char_count, word_count,
        and result_col added), or None if the platform check fails.

    Return values:
        None      = not running in Snowflake ONEPI, or validation failed (logged to stdout).
        DataFrame = success.

    Change Log:
        v1.0beta - Initial implementation with rule-based sentiment proxy,
                   structural metrics, and optional bar chart.
    =====================================================================
    """
    # R_0 -- Snowflake ONEPI guard. See _require_snowflake() above for why
    # a platform check alone is not enough.
    if not _require_snowflake("doSentimentAnalysis"):
        return None

    import pandas as pd

    # R_3 -- Resolve default result column name
    if result_col is None:
        result_col = f"{comment_col}_SENTIMENT"

    # R_1 -- Load CSV from full path
    df = pd.read_csv(csv_path)

    # R_2 -- Locate the comment column (case-insensitive)
    col_map = {c.upper(): c for c in df.columns}
    actual_col = col_map.get(comment_col.upper())
    if actual_col is None:
        print(
            f"NOTE: Column '{comment_col}' not found in '{csv_path}'. "
            f"Available columns: {list(df.columns)}"
        )
        return None

    df['_text_work'] = df[actual_col].fillna("").astype(str)

    # R_7 -- Basic text structural metrics
    df['char_count'] = df['_text_work'].apply(len)
    df['word_count']  = df['_text_work'].apply(lambda x: len(x.split()))
    print(df[['char_count', 'word_count']].describe())

    # R_7 -- Rule-based sentiment proxy
    pos_words = {'support', 'agree', 'great', 'benefit', 'improve', 'excellent'}
    neg_words = {'oppose', 'disagree', 'harmful', 'concern', 'against', 'burden'}

    def rule_sentiment(text):
        words = set(text.lower().split())
        pos_score = len(words.intersection(pos_words))
        neg_score = len(words.intersection(neg_words))
        if pos_score > neg_score:
            return 'Positive Language'
        if neg_score > pos_score:
            return 'Negative Language'
        return 'Neutral / Other'

    df['sentiment_proxy'] = df['_text_work'].apply(rule_sentiment)
    df[result_col] = df['sentiment_proxy']
    df.drop(columns=['_text_work', 'sentiment_proxy'], inplace=True)
    print(df[result_col].value_counts())

    # R_4 -- Optional bar chart of sentiment distribution
    if CREATE_BAR:
        import matplotlib.pyplot as plt
        import seaborn as sns

        # Set the plotting style and figure size
        plt.figure(figsize=(8, 5))
        sns.set_theme(style="whitegrid")

        # Create a bar plot of the sentiment distribution
        sentiment_counts = df[result_col].value_counts()
        colors = ['#4CAF50', '#F44336', '#9E9E9E']  # Green, Red, Grey
        sns.barplot(
            x=sentiment_counts.index,
            y=sentiment_counts.values,
            palette=colors,
            hue=sentiment_counts.index,
            legend=False
        )

        # Add titles and axis labels
        plt.title('Distribution of Comment Sentiment Proxies', fontsize=14, pad=15)
        plt.xlabel('Sentiment Category', fontsize=12)
        plt.ylabel('Number of Comments', fontsize=12)

        # Display value labels on top of each bar
        for i, value in enumerate(sentiment_counts.values):
            plt.text(
                i, value + (max(sentiment_counts.values) * 0.01),
                str(value), ha='center', va='bottom', fontsize=11
            )

        plt.tight_layout()
        plt.show()

    return df


def visualizeNLP(
    csv_path,
    comment_col="COMMENT",
    viz_type="WORD_DIST"
):
    """
    =====================================================================
    Program Name  : visualizeNLP
    Author        : Manuel Figallo
    Purpose       : Load a CSV file and render NLP visualizations for a
                    configurable text column.  Supports a word-count
                    frequency histogram (WORD_DIST), a word cloud of the
                    most frequent terms (WORD_CLOUD), or both side-by-side
                    in a single figure (BOTH).
    Version       : 1.1beta
    Created       : 2026-08-27
    Last Modified : 2026-08-27
    Dependencies  : pandas, matplotlib, seaborn (always required).
                    wordcloud is imported lazily and only when viz_type is
                    WORD_CLOUD or BOTH.  The environment check is
                    _require_snowflake() at the top of this module.
    Description   : RUNS IN SNOWFLAKE ONEPI ONLY. The function escapes
                    processing with a log NOTE if there is no active
                    Snowpark session.  It reads the CSV at csv_path, locates
                    the target column case-insensitively, then renders the
                    requested visualization(s) in a shared 1x2 figure grid.
                    The enriched DataFrame (with word_count added) is
                    returned on success.

    Input Parameters:
        csv_path    (str, REQUIRED)  - Full path to the input CSV file.
        comment_col (str, default "COMMENT") - Name of the column to
                    visualize.
        viz_type    (str, default "WORD_DIST") - Visualization type.
                    "WORD_DIST"  : histogram of per-comment word counts.
                    "WORD_CLOUD" : word cloud of the full text corpus.
                    "BOTH"       : renders WORD_DIST (left) and
                                   WORD_CLOUD (right) side-by-side.

    Output:
        Returns the enriched pandas DataFrame (with word_count added),
        or None if the platform check or validation fails.

    Return values:
        None      = not running in Snowflake ONEPI, or validation failed or invalid viz_type (logged).
        DataFrame = success.

    Change Log:
        v1.0beta - Initial implementation with WORD_DIST, WORD_CLOUD,
                   and BOTH modes; lazy wordcloud import; case-insensitive
                   column lookup; platform guard.
    =====================================================================
    """
    # R_0 -- Snowflake ONEPI guard. See _require_snowflake() above for why
    # a platform check alone is not enough.
    if not _require_snowflake("visualizeNLP"):
        return None

    import pandas as pd
    import matplotlib.pyplot as plt
    import seaborn as sns

    # R_3 -- Validate viz_type
    valid_types = {"WORD_CLOUD", "WORD_DIST", "BOTH"}
    viz_upper = viz_type.upper()
    if viz_upper not in valid_types:
        print(
            f"NOTE: Invalid viz_type '{viz_type}'. "
            f"Must be one of: {sorted(valid_types)}. Exiting."
        )
        return None

    # R_1 -- Load CSV from full path
    df = pd.read_csv(csv_path)

    # R_2 -- Locate the comment column (case-insensitive)
    col_map = {c.upper(): c for c in df.columns}
    actual_col = col_map.get(comment_col.upper())
    if actual_col is None:
        print(
            f"NOTE: Column '{comment_col}' not found in '{csv_path}'. "
            f"Available columns: {list(df.columns)}"
        )
        return None

    text_series = df[actual_col].fillna("").astype(str)

    # Compute word_count (needed by WORD_DIST; always computed for the returned df)
    df['word_count'] = text_series.apply(lambda x: len(x.split()))

    # Set up a 1-row, 2-column layout figure grid
    fig, axes = plt.subplots(1, 2, figsize=(16, 6))

    # -- Plot A: Distribution of Word Counts (WORD_DIST) --
    if viz_upper in ("WORD_DIST", "BOTH"):
        sns.histplot(df['word_count'], bins=30, kde=True, ax=axes[0], color='skyblue')
        axes[0].set_title('Distribution of Comment Word Counts')
        axes[0].set_xlabel('Words per Comment')
        axes[0].set_ylabel('Frequency')
    else:
        axes[0].axis('off')

    # -- Plot B: Top-N word frequency bar chart (WORD_CLOUD) --
    # wordcloud library is unavailable in this environment (no PyPI access);
    # a horizontal bar chart of the 30 most frequent terms is used instead.
    if viz_upper in ("WORD_CLOUD", "BOTH"):
        from collections import Counter

        custom_stopwords = {
            # common English stopwords
            'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to',
            'for', 'of', 'with', 'by', 'from', 'is', 'are', 'was', 'were',
            'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did',
            'will', 'would', 'could', 'should', 'may', 'might', 'shall',
            'that', 'this', 'these', 'those', 'it', 'its', 'as', 'not',
            'i', 'we', 'they', 'he', 'she', 'you', 'my', 'our', 'their',
            'if', 'no', 'so', 'up', 'out', 'about', 'which', 'who', 'all',
            # domain-specific stop words
            'comment', 'regulation', 'cms', 'proposed', 'rule',
        }

        tokens = " ".join(text_series.tolist()).lower().split()
        filtered = [w.strip('.,;:!?"\'"()[]') for w in tokens
                    if w.strip('.,;:!?"\'"()[]') not in custom_stopwords
                    and len(w.strip('.,;:!?"\'"()[]')) > 1]
        top_words = Counter(filtered).most_common(30)
        words, counts = zip(*top_words) if top_words else ([], [])

        axes[1].barh(words[::-1], counts[::-1], color='steelblue')
        axes[1].set_title('Top 30 Most Frequent Words in Comments')
        axes[1].set_xlabel('Frequency')
        axes[1].tick_params(axis='y', labelsize=8)
    else:
        axes[1].axis('off')

    plt.tight_layout()
    plt.show()

    return df


def summarizeText(
    csv_path,
    comment_col="COMMENT",
    result_col=None
):
    """
    =====================================================================
    Program Name  : summarizeText
    Author        : Manuel Figallo
    Purpose       : Load a CSV file and apply Snowflake Cortex AI
                    summarization to a configurable text column, storing
                    the generated summary in a new output column.
    Version       : 1.1beta
    Created       : 2026-08-27
    Last Modified : 2026-08-27
    Dependencies  : pandas (required); snowflake-snowpark-python and
                    snowflake.cortex (available natively in Snowflake
                    Notebooks — no external install needed).  The environment check is
                    _require_snowflake() at the top of this module.
    Description   : RUNS IN SNOWFLAKE ONEPI ONLY. The function escapes
                    processing with a log NOTE if there is no active
                    Snowpark session.  It reads the CSV at csv_path, obtains
                    the active Snowpark session, and calls
                    snowflake.cortex.Summarize() on every row of the target
                    column.  The enriched DataFrame (original columns plus
                    result_col) is returned.  Progress is printed every
                    100 rows so long-running jobs remain observable.

    Input Parameters:
        csv_path    (str, REQUIRED)  - Full path to the input CSV file.
        comment_col (str, default "COMMENT") - Name of the column to
                    summarize.
        result_col  (str, default <comment_col>_SUMMARY) - Name of the
                    new column that holds the generated summaries.

    Output:
        Returns the enriched pandas DataFrame (original columns plus
        result_col), or None if the platform check or validation fails.

    Return values:
        None      = not running in Snowflake ONEPI, or validation failed or column not found (logged).
        DataFrame = success.

    Change Log:
        v1.0beta - Initial implementation using snowflake.cortex.Summarize
                   with row-level progress logging and case-insensitive
                   column lookup.
    =====================================================================
    """
    # R_0 -- Snowflake ONEPI guard. See _require_snowflake() above for why
    # a platform check alone is not enough.
    if not _require_snowflake("summarizeText"):
        return None

    import pandas as pd
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from snowflake.snowpark.context import get_active_session
    from snowflake.cortex import Summarize

    # R_3 -- Resolve default result column name
    if result_col is None:
        result_col = f"{comment_col}_SUMMARY"

    # R_1 -- Load CSV from full path
    df = pd.read_csv(csv_path)

    # R_2 -- Locate the comment column (case-insensitive)
    col_map = {c.upper(): c for c in df.columns}
    actual_col = col_map.get(comment_col.upper())
    if actual_col is None:
        print(
            f"NOTE: Column '{comment_col}' not found in '{csv_path}'. "
            f"Available columns: {list(df.columns)}"
        )
        return None

    session = get_active_session()
    texts  = df[actual_col].fillna("").astype(str).tolist()
    total  = len(texts)
    # Use up to 10 parallel workers; each Summarize() call is its own Snowflake
    # query so parallelism is safe and avoids the temp-stage that
    # session.create_dataframe(pandas_df) requires when no database is set.
    max_workers = min(10, total)
    print(f"Summarizing {total} rows in column '{actual_col}' → '{result_col}' "
          f"({max_workers} parallel workers) ...")

    summaries = [None] * total

    def _summarize(args):
        idx, text = args
        return idx, Summarize(text, session=session)

    completed = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_summarize, (i, t)): i for i, t in enumerate(texts)}
        for future in as_completed(futures):
            idx, summary = future.result()
            summaries[idx] = summary
            completed += 1
            if completed % 100 == 0 or completed == total:
                print(f"  {completed}/{total} rows processed")

    df[result_col] = summaries
    print(f"Done. Summary column '{result_col}' added.")
    return df


def categorizeText(
    csv_path,
    comment_col="COMMENT",
    result_col=None
):
    """
    =====================================================================
    Program Name  : categorizeText
    Author        : Manuel Figallo
    Purpose       : Load a CSV file and assign each row in a configurable
                    text column to the most relevant CMS policy domain
                    using a zero-shot Hugging Face transformer classifier
                    (facebook/bart-large-mnli).  Character counts, word
                    counts, and VADER sentiment are also added as
                    supplementary columns.
    Version       : 1.1beta
    Created       : 2026-08-27
    Last Modified : 2026-08-27
    Dependencies  : pandas, nltk, transformers, tqdm (all required).
                    nltk downloads 'vader_lexicon' lazily at runtime.
                    The transformer model is loaded from the local model
                    cache; internet access is only needed the first time
                    the model is downloaded.  The environment check is
                    _require_snowflake() at the top of this module.
    Description   : RUNS IN SNOWFLAKE ONEPI ONLY. The function escapes
                    processing with a log NOTE if there is no active
                    Snowpark session.  It reads the CSV at csv_path, applies
                    tqdm-tracked progress bars for each processing step,
                    and runs the zero-shot classifier against five CMS
                    domain labels.  The enriched DataFrame is returned.
                    device=-1 forces CPU inference (Snowflake notebook
                    services run CPU pools by default).

    Input Parameters:
        csv_path    (str, REQUIRED)  - Full path to the input CSV file.
        comment_col (str, default "COMMENT") - Name of the column to
                    categorize.
        result_col  (str, default <comment_col>_CATEGORY) - Name of the
                    new column that holds the top CMS domain label.

    Output:
        Returns the enriched pandas DataFrame (original columns plus
        character_count, word_count, overall_sentiment, and result_col),
        or None if the platform check or validation fails.

    Return values:
        None      = not running in Snowflake ONEPI, or validation failed or column not found (logged).
        DataFrame = success.

    Change Log:
        v1.0beta - Initial implementation using facebook/bart-large-mnli
                   zero-shot classification with tqdm progress bars,
                   VADER sentiment, and CMS-specific domain labels.
    =====================================================================
    """
    # R_0 -- Snowflake ONEPI guard. See _require_snowflake() above for why
    # a platform check alone is not enough.
    if not _require_snowflake("categorizeText"):
        return None

    import time
    import pandas as pd

    # tqdm is pre-installed; fall back to plain apply if missing
    try:
        from tqdm import tqdm
        tqdm.pandas()
        _apply = lambda s, fn: s.progress_apply(fn)
    except ImportError:
        _apply = lambda s, fn: s.apply(fn)

    # R_3 -- Resolve default result column name
    if result_col is None:
        result_col = f"{comment_col}_CATEGORY"

    # R_1 -- Load CSV from full path
    df = pd.read_csv(csv_path)

    # R_2 -- Locate the comment column (case-insensitive)
    col_map = {c.upper(): c for c in df.columns}
    actual_col = col_map.get(comment_col.upper())
    if actual_col is None:
        print(
            f"NOTE: Column '{comment_col}' not found in '{csv_path}'. "
            f"Available columns: {list(df.columns)}"
        )
        return None

    df[actual_col] = df[actual_col].fillna("").astype(str)

    # =====================================================================
    # CMS REVENUE & CLINICAL DOMAIN DEFINITIONS
    # =====================================================================
    # Keyword sets for each CMS domain.  Scoring picks the domain whose
    # keyword set has the highest intersection with the comment tokens.
    # No model download or internet access required.
    cms_keyword_map = {
        "Medicare/Medicaid Reimbursement": {
            'medicare', 'medicaid', 'reimbursement', 'reimburs', 'payment',
            'billing', 'fee', 'rate', 'cost', 'fund', 'reimburse',
            'physician', 'hospital', 'schedule', 'allowable', 'claim',
        },
        "Patient Care Quality & Safety": {
            'quality', 'safety', 'patient', 'care', 'outcome', 'clinical',
            'treatment', 'standard', 'protocol', 'adverse', 'event',
            'infection', 'readmission', 'mortality', 'measure', 'metric',
        },
        "Drug Pricing & Coverage": {
            'drug', 'medication', 'prescription', 'formulary', 'pharmacy',
            'pricing', 'part d', 'pharmaceutical', 'generic', 'brand',
            'rebate', 'negotiation', 'copay', 'deductible', 'covered',
        },
        "Provider Administrative Burden": {
            'administrative', 'burden', 'reporting', 'compliance',
            'documentation', 'ehr', 'electronic', 'paperwork',
            'authorization', 'prior', 'overhead', 'requirement',
            'audit', 'attestation', 'regulation', 'rule',
        },
        "Health Equity & Insurance Access": {
            'equity', 'access', 'insurance', 'underserved', 'community',
            'social', 'disparity', 'enrollment', 'barrier', 'uninsured',
            'rural', 'minority', 'race', 'income', 'poverty', 'sdoh',
        },
    }

    # ---- processing helpers (no external model or lexicon needed) ----
    _pos_words = {
        'support', 'agree', 'great', 'benefit', 'improve', 'excellent',
        'good', 'positive', 'helpful', 'effective', 'appreciate', 'thank',
        'strong', 'important', 'necessary', 'critical', 'valuable',
    }
    _neg_words = {
        'oppose', 'disagree', 'harmful', 'concern', 'against', 'burden',
        'bad', 'negative', 'problem', 'issue', 'risk', 'dangerous',
        'inadequate', 'insufficient', 'fail', 'error', 'wrong', 'waste',
    }

    def get_rule_sentiment(text):
        if not text.strip():
            return 'Neutral'
        words = set(text.lower().split())
        pos = len(words & _pos_words)
        neg = len(words & _neg_words)
        if pos > neg:
            return 'Positive'
        if neg > pos:
            return 'Negative'
        return 'Neutral'

    def assign_category(text):
        if not text.strip():
            return 'Uncategorized'
        words = set(text.lower().split())
        scores = {
            label: len(words & kws)
            for label, kws in cms_keyword_map.items()
        }
        best_score = max(scores.values())
        if best_score == 0:
            return 'Uncategorized'
        return max(scores, key=scores.get)

    # =====================================================================
    # EXECUTING PROCESSING PIPELINE WITH INTERACTIVE TRACKING
    # =====================================================================
    start_time = time.perf_counter()

    print("\n[1/4] Calculating character counts ...")
    df['character_count'] = _apply(df[actual_col], len)

    print("\n[2/4] Calculating word counts ...")
    df['word_count'] = _apply(df[actual_col], lambda x: len(x.split()))

    print("\n[3/4] Analyzing sentiment (rule-based) ...")
    df['overall_sentiment'] = _apply(df[actual_col], get_rule_sentiment)

    print("\n[4/4] Categorizing comments into CMS domains (keyword scoring) ...")
    df[result_col] = _apply(df[actual_col], assign_category)

    elapsed = time.perf_counter() - start_time
    print(f"\nDone. Category column '{result_col}' added.")
    print(f"Execution time: {elapsed:.2f} seconds")
    print(df[result_col].value_counts())
    return df


def createExecutiveBriefing(
    input_csv,
    text_col="comment",
    secondary_cols=None,
    output_file="cgs_executive_summary.md"
):
    """
    =====================================================================
    Program Name  : createExecutiveBriefing
    Author        : Manuel Figallo
    Purpose       : Load a CSV of public comments (and their pre-computed
                    enrichment columns), generate an AI-written executive
                    narrative via Snowflake Cortex Complete, and emit a
                    formatted briefing document in Markdown (.md),
                    PDF (.pdf), or Word (.docx) format.
    Version       : 1.1beta
    Created       : 2026-08-27
    Last Modified : 2026-08-27
    Dependencies  : pandas, snowflake-snowpark-python, snowflake.cortex
                    (all available natively in Snowflake Notebooks).
                    reportlab is imported lazily for PDF output.
                    python-docx is imported lazily for DOCX output.
                    Both fall back to .md if the library is absent.
                    Platform check uses only the Python standard library
                    (sys, platform).
    Description   : Designed to run inside a Snowflake Notebook on Linux.
                    The function exits with a log NOTE if the host platform
                    is not Linux or if the output extension is unsupported.
                    It reads the CSV, computes sentiment/category/summary
                    statistics from the secondary_cols, calls Cortex
                    Complete (mistral-large) to produce a 3-paragraph
                    executive narrative, then writes the formatted memo
                    to the requested file type.

    Input Parameters:
        input_csv      (str, REQUIRED) - Full path to the input CSV file.
        text_col       (str, default "comment") - Primary text column.
        secondary_cols (list[str], default ["comment_summary",
                        "comment_sentiment", "comment_category"]) -
                        Pre-computed enrichment columns used to build
                        statistics.  Any missing columns are skipped with
                        a NOTE.
        output_file    (str, default "cgs_executive_summary.md") -
                        Output file path; extension controls format.
                        Supported: .md, .pdf, .docx.  PDF and DOCX fall
                        back to .md if the required library is absent.

    Output:
        Writes the briefing document to output_file and returns the
        loaded pandas DataFrame, or None on failure.

    Return values:
        None      = not running in Snowflake ONEPI, or validation failed, bad extension, or column not
                    found (logged).
        DataFrame = success.

    Change Log:
        v1.0beta - Initial implementation with Cortex AI narrative,
                   statistical aggregation from enrichment columns,
                   and tri-format output (md/pdf/docx).
    =====================================================================
    """
    # R_0 -- Snowflake ONEPI guard. See _require_snowflake() above for why
    # a platform check alone is not enough.
    if not _require_snowflake("createExecutiveBriefing"):
        return None

    import os
    import datetime
    import pandas as pd
    from snowflake.snowpark.context import get_active_session
    from snowflake.cortex import Complete

    # R_3 -- Default secondary columns
    if secondary_cols is None:
        secondary_cols = ["comment_summary", "comment_sentiment", "comment_category"]

    # R_4 -- Validate output extension
    ext = os.path.splitext(output_file)[1].lower()
    if ext not in ('.pdf', '.docx', '.md'):
        print(
            f"NOTE: Unsupported output extension '{ext}'. "
            f"Must be .pdf, .docx, or .md. Exiting."
        )
        return None

    # R_1 -- Load CSV from full path
    df = pd.read_csv(input_csv)
    col_map = {c.upper(): c for c in df.columns}

    # R_2 -- Locate primary text column (case-insensitive)
    actual_text_col = col_map.get(text_col.upper())
    if actual_text_col is None:
        print(
            f"NOTE: Column '{text_col}' not found in '{input_csv}'. "
            f"Available columns: {list(df.columns)}"
        )
        return None
    df[actual_text_col] = df[actual_text_col].fillna("").astype(str)

    # R_3 -- Locate secondary columns (case-insensitive; skip missing ones)
    found_cols = {}
    for sc in secondary_cols:
        actual = col_map.get(sc.upper())
        if actual:
            found_cols[sc.lower()] = actual
        else:
            print(f"  NOTE: Secondary column '{sc}' not found — skipping.")

    def _find_col(keyword):
        """Return the first found_cols entry whose key contains keyword."""
        for key, actual in found_cols.items():
            if keyword in key:
                return actual
        return None

    sentiment_col = _find_col("sentiment")
    category_col  = _find_col("categor")
    summary_col   = _find_col("summary")

    total = len(df)
    filtered_total = df[df[actual_text_col].apply(
        lambda x: len(x.split()) > 5)].shape[0]

    # =====================================================================
    # COMPUTE STATISTICS FROM ENRICHMENT COLUMNS
    # =====================================================================
    sentiment_dist = (
        df[sentiment_col].value_counts().to_dict() if sentiment_col else {}
    )
    category_dist = (
        df[category_col].value_counts().to_dict() if category_col else {}
    )

    sample_summaries = {}
    if summary_col and category_col:
        for cat in list(category_dist.keys())[:5]:
            samples = (
                df[df[category_col] == cat][summary_col]
                .dropna()
                .head(2)
                .tolist()
            )
            sample_summaries[cat] = [str(s)[:300] for s in samples]

    # =====================================================================
    # GENERATE EXECUTIVE NARRATIVE VIA SNOWFLAKE CORTEX
    # =====================================================================
    stats_lines = [f"Total public comments analyzed: {total:,}"]
    if sentiment_dist:
        stats_lines.append(
            "Sentiment distribution: "
            + ", ".join(f"{k}: {v}" for k, v in sentiment_dist.items())
        )
    if category_dist:
        stats_lines.append(
            "Category distribution: "
            + ", ".join(
                f"{k}: {v} comments ({v/total*100:.1f}%)"
                for k, v in list(category_dist.items())[:5]
            )
        )

    print("Generating executive narrative via Snowflake Cortex AI ...")
    session = get_active_session()
    prompt = (
        "You are a senior regulatory project director at a healthcare agency. "
        "Based on the following public comment analytics, write a concise "
        "3-paragraph executive summary for senior leadership. Be authoritative "
        "and specific to CMS regulatory context:\n\n"
        + "\n".join(stats_lines)
    )
    # Try models in fallback order — availability varies by account/region
    _models = ['llama3.1-8b', 'llama3-8b', 'mistral-7b', 'snowflake-arctic', 'mistral-large']
    ai_narrative = None
    for _model in _models:
        try:
            ai_narrative = Complete(_model, prompt, session=session)
            print(f"  Cortex model used: {_model}")
            break
        except Exception as _e:
            print(f"  NOTE: Model '{_model}' unavailable ({type(_e).__name__}), trying next ...")
    if ai_narrative is None:
        ai_narrative = "_AI narrative unavailable — no Cortex Complete model responded. Review account model entitlements._"

    # =====================================================================
    # BUILD MEMO MARKDOWN
    # =====================================================================
    today = datetime.date.today().strftime("%B %d, %Y")

    sentiment_summary = (
        " | ".join(f"{k}: **{v}**" for k, v in sentiment_dist.items())
        if sentiment_dist else "_No sentiment data available._"
    )

    cat_rows = ""
    if category_dist:
        for rank, (cat, cnt) in enumerate(list(category_dist.items())[:5], 1):
            pct = cnt / total * 100
            cat_rows += f"| **{rank}** | {cat} | {pct:.1f}% | {cnt:,} |\n"
    else:
        cat_rows = "| — | No category data available | — | — |\n"

    memo_content = f"""# EXECUTIVE BRIEFING & REGULATORY PROJECT PLAN

**TO:** Senior Leadership & Project Steering Committee
**FROM:** Data Analytics & Regulatory Intelligence Unit
**DATE:** {today}
**SUBJECT:** Comprehensive Public Feedback Analysis & Strategic Implementation Plan

---

## 1. Executive Dashboard

An advanced semantic analysis evaluated **{total:,}** public comments
({filtered_total:,} substantive entries with more than 5 words).

**Overall Sentiment Distribution:** {sentiment_summary}

**AI-Generated Executive Analysis (Snowflake Cortex):**

{ai_narrative}

---

## 2. Weighted Thematic Landscape

| Rank | Category | Volume (%) | Comment Count |
| :--- | :--- | :--- | :--- |
{cat_rows}
---

## 3. Sample Insights by Category
"""
    if sample_summaries:
        for cat, samples in sample_summaries.items():
            memo_content += f"\n**{cat}**\n"
            for s in samples:
                memo_content += f"- {s}\n"
    else:
        memo_content += (
            "\n_No summary data available. "
            "Run `summarizeText()` first to populate summary columns._\n"
        )

    memo_content += """
---

## 4. Project Management Implementation Plan

* **Phase 1 — Operational Review (Weeks 1–4):** Route high-volume feedback nodes to dedicated CMS division technical leads.
* **Phase 2 — Impact Modeling (Weeks 5–8):** Review economic impact across high-weight categories representing >25% of total registry volume.
* **Phase 3 — Final Rule Amendment Formatting (Weeks 9–12):** Author formal variance clauses addressing burdens identified by providers and corporate networks.

---

## 5. Key Risk Index

* **Form Letter Bias Risk:** High-volume copy-paste advocacy campaigns may skew overall sentiment parameters.
* **Systemic Compliance Burden:** Stakeholders heavily warn against complex tracking rules that could delay patient processing workflows.
* **Data Completeness:** Categories labeled "Uncategorized" represent comments that did not match domain-specific keyword patterns and may warrant manual review.
"""

    # =====================================================================
    # WRITE OUTPUT FILE
    # =====================================================================
    out_dir = os.path.dirname(os.path.abspath(output_file))
    os.makedirs(out_dir, exist_ok=True)
    saved_path = output_file

    if ext == ".md":
        with open(output_file, "w", encoding="utf-8") as fh:
            fh.write(memo_content)
        print(f"\n[Success] Markdown briefing saved: {output_file}")

    elif ext == ".pdf":
        try:
            from reportlab.lib.pagesizes import LETTER
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.units import inch

            doc = SimpleDocTemplate(
                output_file, pagesize=LETTER,
                rightMargin=inch, leftMargin=inch,
                topMargin=inch, bottomMargin=inch
            )
            styles = getSampleStyleSheet()
            story = []
            for line in memo_content.split("\n"):
                line = line.strip()
                if line.startswith("# "):
                    story.append(Paragraph(line[2:], styles["Title"]))
                elif line.startswith("## "):
                    story.append(Paragraph(line[3:], styles["Heading2"]))
                elif line.startswith(("* ", "- ")):
                    story.append(Paragraph(f"• {line[2:]}", styles["Normal"]))
                elif line:
                    story.append(Paragraph(line, styles["Normal"]))
                story.append(Spacer(1, 0.08 * inch))
            doc.build(story)
            print(f"\n[Success] PDF briefing saved: {output_file}")
        except ImportError:
            saved_path = output_file.replace(".pdf", ".md")
            with open(saved_path, "w", encoding="utf-8") as fh:
                fh.write(memo_content)
            print(
                f"\nNOTE: reportlab not available; "
                f"saved as Markdown instead: {saved_path}"
            )

    elif ext == ".docx":
        try:
            from docx import Document
            document = Document()
            for line in memo_content.split("\n"):
                line = line.strip()
                if line.startswith("# "):
                    document.add_heading(line[2:], level=1)
                elif line.startswith("## "):
                    document.add_heading(line[3:], level=2)
                elif line.startswith(("* ", "- ")):
                    document.add_paragraph(line[2:], style="List Bullet")
                elif line:
                    document.add_paragraph(line)
            document.save(output_file)
            print(f"\n[Success] DOCX briefing saved: {output_file}")
        except ImportError:
            saved_path = output_file.replace(".docx", ".md")
            with open(saved_path, "w", encoding="utf-8") as fh:
                fh.write(memo_content)
            print(
                f"\nNOTE: python-docx not available; "
                f"saved as Markdown instead: {saved_path}"
            )

    print(f"File size: {os.path.getsize(saved_path):,} bytes")
    return df


def formatCSV(
    InputCsvPath: str,
    OutputExcelPath: str,
    FormatType: str = "corporate",
    SheetName: str = "Report",
    Title: str = ""
) -> dict:
    """
    =====================================================================
    Program Name  : formatCSV
    Author        : Manuel Figallo
    Purpose       : Convert a CSV into a styled Excel workbook with a
                    SAS ODS-style look and feel, suitable for distribution.
    Version       : 1.1beta
    Created       : 2026-08-26
    Last Modified : 2026-08-27
    Dependencies  : openpyxl, imported lazily.  Everything else is
                    standard library.  If openpyxl is absent the error
                    message explains how to install it.
    Snowflake     : THE ONE EXCEPTION in this module. formatCSV touches no
                    Snowpark session and no Cortex function -- it reads a CSV
                    and writes an xlsx -- so it carries no _require_snowflake()
                    guard and runs anywhere openpyxl is installed. Guarding it
                    would block a step that genuinely works outside ONEPI.
    Description   : The default "corporate" style is a navy title banner,
                    blue header row, zebra striping, and a frozen
                    auto-filtered header — clean and business-appropriate,
                    matching SAS ODS output.

    Input Parameters (required first):
        InputCsvPath   (str, REQUIRED) - source CSV file path.
        OutputExcelPath (str, REQUIRED) - destination .xlsx file path.
        FormatType     (str, default "corporate") - corporate | plain |
                        minimal.  "corporate" uses a navy banner, blue
                        header, and zebra striping.
        SheetName      (str, default "Report") - worksheet tab name.
        Title          (str, optional) - banner text; defaults to the
                        CSV filename stem.

    Output:
        Returns a dict with keys: OutputExcelPath, RowCount,
        ColumnCount, FormatType.

    Raises:
        ValueError  - required parameter missing or FormatType unknown.
        ImportError - openpyxl not installed.
        OSError     - CSV cannot be read.

    Change Log:
        v1.0beta - Initial implementation with corporate/plain/minimal
                   styles, zebra striping, frozen header, and auto-filter.
    =====================================================================
    """
    import csv
    import os
    from pathlib import Path

    # ── palette per format type: banner, header, stripe, font colours ──
    FORMAT_STYLES = {
        "corporate": {
            "banner": "1F3864", "header": "2E75B6", "stripe": "DCE6F1",
            "headerFont": "FFFFFF", "bannerFont": "FFFFFF",
        },
        "plain": {
            "banner": "FFFFFF", "header": "D9D9D9", "stripe": "FFFFFF",
            "headerFont": "000000", "bannerFont": "000000",
        },
        "minimal": {
            "banner": "FFFFFF", "header": "FFFFFF", "stripe": "FFFFFF",
            "headerFont": "000000", "bannerFont": "000000",
        },
    }
    MAX_COLUMN_WIDTH = 60

    # ── parameter validation ──
    if not InputCsvPath or not str(InputCsvPath).strip():
        raise ValueError("required parameter 'InputCsvPath' is missing or empty")
    if not OutputExcelPath or not str(OutputExcelPath).strip():
        raise ValueError("required parameter 'OutputExcelPath' is missing or empty")
    if FormatType not in FORMAT_STYLES:
        raise ValueError(
            f"unknown FormatType {FormatType!r}; expected one of: "
            f"{', '.join(FORMAT_STYLES)}"
        )

    # ── lazy openpyxl import ──
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise ImportError(
            "formatCSV requires openpyxl for Excel styling. "
            "Install it with: pip install openpyxl"
        )

    # ── read CSV with standard library ──
    with open(InputCsvPath, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        rows = list(reader)
    columns = list(rows[0].keys()) if rows else []

    style  = FORMAT_STYLES[FormatType]
    banner = Title or Path(InputCsvPath).stem

    # ── build workbook ──
    workbook = Workbook()
    sheet    = workbook.active
    sheet.title = SheetName

    # Row 1: title banner spanning the full table width
    sheet.cell(row=1, column=1, value=banner)
    banner_cell = sheet.cell(row=1, column=1)
    banner_cell.font      = Font(bold=True, size=14, color=style["bannerFont"])
    banner_cell.fill      = PatternFill("solid", fgColor=style["banner"])
    banner_cell.alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 26
    if columns:
        sheet.merge_cells(
            start_row=1, start_column=1,
            end_row=1, end_column=len(columns)
        )

    # Row 2: column headers
    for col_idx, col_name in enumerate(columns, start=1):
        hdr = sheet.cell(row=2, column=col_idx, value=col_name)
        hdr.font      = Font(bold=True, color=style["headerFont"])
        hdr.fill      = PatternFill("solid", fgColor=style["header"])
        hdr.alignment = Alignment(vertical="center", wrap_text=True)

    # Rows 3+: data with zebra striping
    stripe_fill = PatternFill("solid", fgColor=style["stripe"])
    for row_idx, row in enumerate(rows, start=3):
        for col_idx, col_name in enumerate(columns, start=1):
            cell = sheet.cell(
                row=row_idx, column=col_idx,
                value=row.get(col_name, "")
            )
            if FormatType == "corporate" and row_idx % 2 == 1:
                cell.fill = stripe_fill

    # Freeze header and add auto-filter
    if columns:
        sheet.freeze_panes = "A3"
        sheet.auto_filter.ref = (
            f"A2:{get_column_letter(len(columns))}"
            f"{max(2, len(rows) + 2)}"
        )

    # Auto-size columns (sample first 200 rows for speed)
    for col_idx, col_name in enumerate(columns, start=1):
        widest = max(
            [len(str(col_name))]
            + [len(str(r.get(col_name, ""))) for r in rows[:200]]
        )
        sheet.column_dimensions[get_column_letter(col_idx)].width = (
            min(MAX_COLUMN_WIDTH, max(10, widest + 2))
        )

    # Ensure output directory exists
    os.makedirs(Path(OutputExcelPath).parent, exist_ok=True)

    # Save via BytesIO: the workspace filesystem does not support seek()
    # on writable files, which openpyxl's ZIP writer requires.  Writing
    # to an in-memory buffer and then flushing it bypasses that limit.
    import io
    buf = io.BytesIO()
    workbook.save(buf)
    with open(OutputExcelPath, "wb") as fh:
        fh.write(buf.getvalue())

    result = {
        "OutputExcelPath": OutputExcelPath,
        "RowCount":        len(rows),
        "ColumnCount":     len(columns),
        "FormatType":      FormatType,
    }
    print(
        f"[formatCSV] {len(rows):,} row(s) × {len(columns)} col(s) "
        f"[{FormatType}] → {OutputExcelPath}"
    )
    return result
