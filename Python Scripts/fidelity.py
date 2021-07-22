import numpy as np


polarization = np.radians(30)
phaseDelay = np.radians(60)
EZ = 0.099430222
EX = 0.132862352
EY = 0.157386698

I = np.matrix([[1, 0],
               [0, 1]])

Z = np.matrix([[1, 0],
               [0, -1]])

X = np.matrix([[0, 1],
               [1, 0]])

Y = np.matrix([[0, complex(0, -1)],
               [complex(0, 1), 0]])

actual = 0.5 * (1 * I + EZ * Z + EX * X + EY * Y)

# thetaHWP = np.radians(input[0])
# HWP = np.matrix([[np.cos(2 * thetaHWP), np.sin(2 * thetaHWP)],
#                  [np.sin(2 * thetaHWP), -np.cos(2 * thetaHWP)]])

# thetaQWP = np.radians(input[1])
# QWP = np.matrix([[np.square(np.cos(thetaQWP)) + complex(0, 1) * np.square(np.sin(thetaQWP)), complex(1, -1) * (np.cos(thetaQWP) * np.sin(thetaQWP))],
#                  [complex(1, -1) * (np.cos(thetaQWP) * np.sin(thetaQWP)), np.square(np.sin(thetaQWP)) + complex(0, 1) * np.square(np.cos(thetaQWP))]])

psi = np.matrix([[np.cos(polarization)],
                 [np.sin(polarization) * complex(np.cos(phaseDelay), np.sin(phaseDelay))]])

fidelity = psi.getH() * actual * psi

print(fidelity.real.item(0, 0))
