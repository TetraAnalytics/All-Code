def severity_calculation(ltv):
    if ltv >= 0.7:
        return 0.6
    elif ltv >= 0.6:
        return 0.4
    else:
        return 0.2