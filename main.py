import math
import openpyxl
import matplotlib.pyplot as plt
book = openpyxl.open("logs.xlsx", read_only=True)
sheet = book.active


def calculatedEkEg(phiLA, lyambdaLA, Hla):
    # перевод радианы в градусы и обратно
    grad_to_rad = math.pi / 180
    rad_to_grad = 180 / math.pi
    # пересчет широты и долготы ЛА в радианы
    phiLA = phiLA * grad_to_rad
    lyambdaLA = lyambdaLA * grad_to_rad

    # иницилизация начальных данных
    R = 6371000
    phit1 = 55.394985533877104
    phit1 = phit1 * grad_to_rad
    lyambdat1 = 37.941730196345475
    lyambdat1 = lyambdat1 * grad_to_rad
    phit2 = 55.402753091364076
    phit2 = phit2 * grad_to_rad
    lyambdat2 = 37.93244510028853
    lyambdat2 = lyambdat2 * grad_to_rad
    Vyd = 1.3
    Vp = 75
    Ht2 = 179
    Ht1 = 179
    deltaDkrm = 1000
    deltaDtp = 100
    alpha0 = 4
    alpha0 = alpha0 * grad_to_rad

    list_Ekv = []
    list_Egv = []
    # расчет при H больше 60 м от впп, угол наклона траектории не меняется
    if Hla - Ht2 > 60:
        # параметры для расчета курса и длины ВПП
        delta_lyambdaA = (phit2 - phit1) * R
        delta_phiA = (lyambdat2 - lyambdat1) * R * math.cos(phit1)

        # Расчет курса и длины ВПП
        psi_vpp = math.atan(delta_lyambdaA / delta_phiA)
        deltaDvpp = (delta_phiA ** 2 + delta_lyambdaA ** 2) ** 0.5

        # параметры для расчета горизонтальной дальности
        delta_phi1 = (phit1 - phiLA) * R
        delta_lyambda1 = (lyambdat1 - lyambdaLA) * R * math.cos(phit1)

        # горизонтальная дальность до ближнего торца ВПП
        Dbtvpp = (delta_phi1 ** 2 + delta_lyambda1 ** 2) ** 0.5

        # параметры для расчета пеленга
        delta_phi2 = (phit2 - phiLA) * R + deltaDkrm * math.cos(psi_vpp)
        delta_lyambda2 = (lyambdat2 - lyambdaLA) * R * math.cos(phit2) + deltaDkrm * math.cos(psi_vpp)

        # расчет пеленга и горизонтальной дальности до ВКГРМ:
        Pvkgrm = math.atan(delta_lyambda2 / delta_phi2)
        Dvkgrm = (delta_phi2 ** 2 + delta_lyambda2 ** 2) ** 0.5

        # расчетные параметры для угла места ВКГРМ
        Hvkgrm = Ht2 - (deltaDvpp - deltaDtp) * math.tan(alpha0)
        # угол места ВКГРМ:
        alphavkgrm = math.atan((Hla - Hvkgrm) / Dvkgrm)

        # угловые отклонения ЛА аппарата от заданной траектории посадки:
        Ekv = Pvkgrm - psi_vpp
        Egv = alphavkgrm - alpha0
    #    расчет при H ЛА меньше 60, угол наклона траектории уменьшается (процесс выравнивания).
    else:
        # начальная высота высота траектории и длина впп
        H0 = 10
        deltaDvpp = 100

        # параметры для расчета горизонтальной дальности
        delta_phi1 = (phit1 - phiLA) * R
        delta_lyambda1 = (lyambdat1 - lyambdaLA) * R * math.cos(phit1)
        delta_lyambdaA = (phit2 - phit1) * R
        delta_phiA = (lyambdat2 - lyambdat1) * R * math.cos(phit1)

        # Расчет курса и длины ВПП
        psi_vpp = math.atan(delta_lyambdaA / delta_phiA)
        # горизонтальная дальность до ближнего торца ВПП
        Dbtvpp = (delta_phi1 ** 2 + delta_lyambda1 ** 2) ** 0.5

        delta_phi2 = (phit2 - phiLA) * R
        delta_lyambda2 = (lyambdat2 - lyambdaLA) * R * math.cos(phit2)

        Pvkgrm = math.atan(delta_lyambda2 / delta_phi2)
        Dvkgrm = (delta_phi2 ** 2 + delta_lyambda2 ** 2) ** 0.5
        delta_Hvkgrm = Dbtvpp * deltaDvpp * (math.tan(alpha0 / H0))**2
        if delta_Hvkgrm < deltaDvpp * Vyd / Vp:
            delta_Hvkgrm = deltaDvpp * Vyd / Vp
        Hvkgrm = Ht2 - delta_Hvkgrm

        # угол места ВКГРМ
        alphavkgrm = math.atan((Hla - Hvkgrm) / Dvkgrm)

    #   текущий угор наклона заданной траектории посадки
        alpha = math.atan(delta_Hvkgrm/deltaDvpp)

        Ekv = Pvkgrm - psi_vpp
        Egv = alphavkgrm - alpha
    list_Ekv.append(Ekv * rad_to_grad)
    list_Egv.append(Egv * rad_to_grad)

    print(list_Ekv, list_Egv)
    return list_Ekv, list_Egv



# Считывания координат и высоты ЛА
for row in range(1, sheet.max_row):
    phiLA = float(sheet[row+1][0].value)
    lyambdaLA = float(sheet[row+1][1].value)
    Hla = float(sheet[row+1][2].value)
    # условие завершения, так как система первой категории

    print(phiLA, lyambdaLA, Hla)
    calculatedEkEg(phiLA, lyambdaLA, Hla)
