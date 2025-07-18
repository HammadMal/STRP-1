
def calculate_clo_scores(clo_assessments, student_scores):
    result = {}

    for student_id in student_scores:
        result[student_id] = {}

        for clo, modules in clo_assessments.items():
            total_weight = sum(item['weight'] for item in modules)
            weighted_score = 0

            for item in modules:
                module = item['module']
                max_score = item['max_score']
                weight = item['weight']
                try:
                    score = float(student_scores[student_id].get(module, 0))
                except (ValueError, TypeError):
                    score = 0.0

                if max_score and max_score != 0:
                    normalized = (score / max_score) * weight
                else:
                    normalized = 0.0

                weighted_score += normalized

            # Always assign a score for the CLO, even if total_weight is 0
            if total_weight != 0:
                final_score = round((weighted_score / total_weight) * 100, 2)
            else:
                final_score = 0.0

            result[student_id][clo] = final_score  # ✅ Ensure every CLO gets an entry

    return result


def calculate_plo_scores(clo_scores, clo_to_plo):
    result = {}
    for student_id, clo_vals in clo_scores.items():
        plo_map = {}
        for clo, score in clo_vals.items():
            mapping = clo_to_plo.get(clo)
            if mapping:
                plo = mapping["PLO"]
                weight = mapping["weight"]
                if plo not in plo_map:
                    plo_map[plo] = {"sum": 0, "total_weight": 0}
                plo_map[plo]["sum"] += score * weight
                plo_map[plo]["total_weight"] += weight
        result[student_id] = {
            plo: round(vals["sum"] / vals["total_weight"], 2)
            for plo, vals in plo_map.items()
        }
    return result

def calculate_grades(clo_assessments, student_scores):
    result = {}
    all_modules = []

    # Flatten all module entries from all CLOs
    for clo_modules in clo_assessments.values():
        all_modules.extend(clo_modules)

    total_weight = sum(item['weight'] for item in all_modules)

    for student_id in student_scores:
        weighted_score = 0
        for item in all_modules:
            module = item['module']
            max_score = item['max_score']
            weight = item['weight']
            try:
                score = float(student_scores[student_id].get(module, 0))
            except (ValueError, TypeError):
                score = 0.0

            normalized = (score / max_score) * weight if max_score else 0.0
            weighted_score += normalized

        final_score = round((weighted_score / total_weight) * 100, 2) if total_weight else 0.0
        result[student_id] = final_score

    return result

def get_letter_grade(percentage):
    if percentage >= 95:
        return "A+"
    elif percentage >= 90:
        return "A"
    elif percentage >= 85:
        return "A-"
    elif percentage >= 80:
        return "B+"
    elif percentage >= 75:
        return "B"
    elif percentage >= 70:
        return "B-"
    elif percentage >= 67:
        return "C+"
    elif percentage >= 63:
        return "C"
    elif percentage >= 60:
        return "C-"
    else:
        return "F"


def get_total_clo_weights(clo_assessments):
    return {
        clo: round(sum(item['weight'] for item in modules), 2)
        for clo, modules in clo_assessments.items()
    }