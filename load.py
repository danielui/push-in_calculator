def calculate_load(P, Ms, Mg, d, ln, mi, loadType):

    if loadType == "axial":
        pmin = (1.7 * P) / (mi * 3.14 * ln * d)
        return pmin

    elif loadType == "twisting moment":
        pmin = (3.4 * Ms) / (mi * 3.14 * ln * (d**2))
        return pmin

    elif loadType == "axial and twisting moment":
        pmin = ((P**2+(2*Ms/d)**2)**0.5) / (mi * 3.14 * ln * d)
        return pmin

    elif loadType == "bending moment":
        pmin = (3 * Mg) / ((ln**2) * (d**2) * 0.7)
        return pmin
    else:
        return 0
