import Levenshtein
import jiwer

def calculate_edit_distance_metrics(text1, text2):
    """
    Calculates Character-level Levenshtein Distance and Word Error Rate (WER).
    """
    # 1. Levenshtein Distance (Character-level)
    char_distance = Levenshtein.distance(text1, text2)
    char_ratio = Levenshtein.ratio(text1, text2)
    
    # 2. Word-Level Metrics (Using the updated jiwer API)
    # process_words calculates WER and extracts alignment stats in one efficient step
    wer_measures = jiwer.process_words(text1, text2)
    
    return {
        "character_level": {
            "levenshtein_distance": char_distance,
            "similarity_ratio": char_ratio
        },
        "word_level": {
            "word_error_rate": wer_measures.wer,
            "substitutions": wer_measures.substitutions,
            "insertions": wer_measures.insertions,
            "deletions": wer_measures.deletions
        }
    }

# --- Example Usage ---
if __name__ == "__main__":
    # Example texts with a few structural differences
    reference_text = "The system shall generate a failure report."
    candidate_text = "The system will generate failure reports."
    reference_text = "2 Abstract . To address growing system complexity in the aerospace industry, a Model-Based Systems Engineering (MBSE) approach has been increasingly adopted for the development of aircraft systems. This calls for a corresponding approach for performing safety assessment to maintain consistency between the system and safety domains. One of the key safety assessment processes for aircraft development is the Functional Hazard Assessment (FHA). The purpose of this paper is to build upon previously published works and introduce the MBFHA framework which describes the language, method, and tool needed for implementing a model-based approach to performing FHA and integrating it into MBSE activities. A customised FHA profile is introduced for the modelling language, an overall workflow along with processes for FHA report and safety requirements generation is presented for the method, and a list of tool constraints is provided. A proof-of-concept is subsequently presented using safety data for the landing gear extension and retraction system of a generic business aircraft.1 Due to the increasing complexity of modern systems in recent years, there has been a growing interest and adoption of a Model-Based Systems Engineering (MBSE) approach to developing aircraft and aircraft systems. Nevertheless, it is important to recognise that for the development of such safety critical systems, system development and safety assessment go hand in hand. Any malfunction can risk the aircraft entering a hazardous situation or even lead to the loss of life. Therefore, as specified by aerospace standards SAE ARP4754A and ARP4761 (SAE International 1996, 2010), it is mandatory to perform safety assessment activities, beginning with a functional hazard assessment (FHA), iteratively during system development. However, while model-based methods for system development are gaining popularity, safety assessment activities are largely still document based, and therefore not linked to MBSE methods or artefacts. In particular, for the generation of an FHA report, data must first be manually extracted from the system model and delivered to the safety team in order for safety experts to perform hazard analysis. With the system model evolving constantly, safety analysis is often performed on outdated information, thus leading to inconsistencies between the system and safety domains (Lai et al. 2021). Moreover, the"
    candidate_text = "Develop a customised FHA profile, define a workflow and processes for FHA reporting and safety requirement generation, specify tool constraints and demonstrate the approach with a proof\u2011of\u2011concept on a generic business aircraft landing gear system"


    print("--- Text Comparison ---")
    print(f"Reference: '{reference_text}'")
    print(f"Candidate: '{candidate_text}'\n")

    metrics = calculate_edit_distance_metrics(reference_text, candidate_text)

    print("1. Character-Level Metrics (Levenshtein):")
    print(f"   - Raw Edit Distance: {metrics['character_level']['levenshtein_distance']} edits")
    print(f"   - Similarity Ratio:  {metrics['character_level']['similarity_ratio']:.4f} (normalized)")
    
    print("\n2. Word-Level Metrics (WER):")
    # WER can be > 1.0 if the candidate has significantly more words than the reference
    print(f"   - Word Error Rate (WER): {metrics['word_level']['word_error_rate']:.4f}")
    print(f"   - Edit Breakdown -> Substitutions: {metrics['word_level']['substitutions']}, "
          f"Insertions: {metrics['word_level']['insertions']}, "
          f"Deletions: {metrics['word_level']['deletions']}")