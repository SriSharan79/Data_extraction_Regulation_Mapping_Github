import nltk
from nltk.translate.bleu_score import sentence_bleu, SmoothingFunction
from rouge_score import rouge_scorer

# Ensure you have the NLTK tokenizer downloaded (required for BLEU tokenization)
nltk.download('punkt', quiet=True)

def calculate_jaccard_similarity(text1, text2):
    """
    Calculates Jaccard Similarity based on unique words (tokens).
    """
    # Simple whitespace/lowercasing tokenization (ignoring punctuation for simplicity)
    words1 = set(text1.lower().split())
    words2 = set(text2.lower().split())
    
    intersection = words1.intersection(words2)
    union = words1.union(words2)
    
    if not union:
        return 0.0
    return len(intersection) / len(union)


def calculate_rouge_scores(reference, candidate):
    """
    Calculates ROUGE-1, ROUGE-2, and ROUGE-L F1-scores.
    """
    scorer = rouge_scorer.RougeScorer(['rouge1', 'rouge2', 'rougeL'], use_stemmer=True)
    scores = scorer.score(reference, candidate)
    
    # Extracting F1-scores for ease of reading
    return {
        "ROUGE-1": scores['rouge1'].fmeasure,
        "ROUGE-2": scores['rouge2'].fmeasure,
        "ROUGE-L": scores['rougeL'].fmeasure
    }


def calculate_bleu_score(reference, candidate):
    """
    Calculates the BLEU score using NLTK.
    """
    # BLEU requires tokenized inputs (lists of words)
    # The reference must be a list of lists (since there can be multiple reference texts)
    ref_tokens = [nltk.word_tokenize(reference.lower())]
    cand_tokens = nltk.word_tokenize(candidate.lower())
    
    # Using a smoothing function to prevent 0 scores for short texts with missing n-grams
    smooth_fn = SmoothingFunction().method1
    
    bleu = sentence_bleu(ref_tokens, cand_tokens, smoothing_function=smooth_fn)
    return bleu


# --- Example Usage ---
if __name__ == "__main__":
    # Define two pieces of text to compare
    text_a = "2 Abstract . To address growing system complexity in the aerospace industry, a Model-Based Systems Engineering (MBSE) approach has been increasingly adopted for the development of aircraft systems. This calls for a corresponding approach for performing safety assessment to maintain consistency between the system and safety domains. One of the key safety assessment processes for aircraft development is the Functional Hazard Assessment (FHA). The purpose of this paper is to build upon previously published works and introduce the MBFHA framework which describes the language, method, and tool needed for implementing a model-based approach to performing FHA and integrating it into MBSE activities. A customised FHA profile is introduced for the modelling language, an overall workflow along with processes for FHA report and safety requirements generation is presented for the method, and a list of tool constraints is provided. A proof-of-concept is subsequently presented using safety data for the landing gear extension and retraction system of a generic business aircraft.1 Due to the increasing complexity of modern systems in recent years, there has been a growing interest and adoption of a Model-Based Systems Engineering (MBSE) approach to developing aircraft and aircraft systems. Nevertheless, it is important to recognise that for the development of such safety critical systems, system development and safety assessment go hand in hand. Any malfunction can risk the aircraft entering a hazardous situation or even lead to the loss of life. Therefore, as specified by aerospace standards SAE ARP4754A and ARP4761 (SAE International 1996, 2010), it is mandatory to perform safety assessment activities, beginning with a functional hazard assessment (FHA), iteratively during system development. However, while model-based methods for system development are gaining popularity, safety assessment activities are largely still document based, and therefore not linked to MBSE methods or artefacts. In particular, for the generation of an FHA report, data must first be manually extracted from the system model and delivered to the safety team in order for safety experts to perform hazard analysis. With the system model evolving constantly, safety analysis is often performed on outdated information, thus leading to inconsistencies between the system and safety domains (Lai et al. 2021). Moreover, the"
    text_b = "SAE ARP4754A"

    print("--- Text Comparison ---")
    print(f"Text 1: '{text_a}'")
    print(f"Text 2: '{text_b}'\n")

    # 1. Jaccard Similarity
    jaccard = calculate_jaccard_similarity(text_a, text_b)
    print(f"1. Jaccard Similarity: {jaccard:.4f}")

    # 2. ROUGE Scores
    # Note: ROUGE typically treats text_a as the 'ground truth/reference'
    rouge = calculate_rouge_scores(reference=text_a, candidate=text_b)
    print("2. ROUGE F1-Scores:")
    for metric, score in rouge.items():
        print(f"   - {metric}: {score:.4f}")

    # 3. BLEU Score
    # Note: BLEU treats text_a as the 'reference' and text_b as the 'candidate translation'
    bleu = calculate_bleu_score(reference=text_a, candidate=text_b)
    print(f"3. BLEU Score: {bleu:.4f}")