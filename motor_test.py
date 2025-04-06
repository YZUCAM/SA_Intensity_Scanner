import thorlabs_apt as apt

if __name__ == '__main__':

    motor = apt.Motor(83844170)
    p1 = motor.position
    print(p1)

    motor.close()
    # motor.move_to(value)
    # motor.move_by(value_step)

