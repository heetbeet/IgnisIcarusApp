import minimalmodbus
import serial
import itertools

minimalmodbus.CLOSE_PORT_AFTER_EACH_CALL = False

def get_instruments(device_ids):
    instances = {}

    device_ids_to_go = list(device_ids)
    modes = ["rtu", "ascii"]
    comrange = [1,2,3,4,5,6,7,8,9,10]
    registers = [512, 320, 0x100, 30002]

    for mode, com, id, reg in itertools.product(modes,
                                                comrange,
                                                device_ids,
                                                registers):
        if id not in device_ids_to_go:
            continue

        comname = f"COM{com}"

        try:
            dev = minimalmodbus.Instrument(comname, id, mode=mode)
            dev.serial.baudrate = 9600
        except serial.serialutil.SerialException:
            continue

        cont = False
        for f in (dev.read_string, dev.read_bits):
            try:
                f(reg, 1)
                print(f"Found dev={id} on {comname}")
                instances[id] = (comname, dev)
                device_ids_to_go.remove(id)
                break

            except (minimalmodbus.NoResponseError, minimalmodbus.InvalidResponseError, minimalmodbus.IllegalRequestError):
                pass

    not_found = set(device_ids).difference(instances)
    if not_found:
        ln = '\n'
        raise ConnectionError(f"Could not connect to Devices: {list(not_found)}, but did find: \n{ln.join([str(i) for i in instances.values()])}")

    return instances


if __name__ == "__main__":
    instruments = get_instruments([1,2,3,4,5,6])
