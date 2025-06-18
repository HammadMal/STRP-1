
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
                except ValueError:
                    score = 0.0
                normalized = (score / max_score) * weight if max_score else 0
                weighted_score += normalized
            result[student_id][clo] = round((weighted_score / total_weight) * 100, 2)
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
