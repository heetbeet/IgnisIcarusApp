import minimalmodbus
import serial
import itertools

minimalmodbus.CLOSE_PORT_AFTER_EACH_CALL = False

def get_instruments(device_ids):
    instances = {}

    """
    instrument.serial.port                     # this is the serial port name
    instrument.serial.baudrate = 19200         # Baud
    instrument.serial.bytesize = 8
    instrument.serial.parity   = serial.PARITY_NONE
    instrument.serial.stopbits = 1
    instrument.serial.timeout  = 0.05          # seconds
    
    instrument.address                         # this is the slave address number
    instrument.mode = minimalmodbus.MODE_RTU   # rtu or ascii mode
    instrument.clear_buffers_before_each_transaction = True
    """

    device_ids_to_go = list(device_ids)

    # Most popular things to iterate over must come last. Rare things must come first
    parities = [serial.PARITY_NONE]   # serial.PARITY_EVEN, serial.PARITY_ODD
    stopbits_ = [serial.STOPBITS_ONE] # serial.STOPBITS_TWO
    bytesizes = [8]                   # 7
    modes = ["rtu", "ascii"]
    comrange = [1,2,3,4,5,6,7,8,9,10]
    registers = [512, 320, 0x100, 1]

    search_count = 0
    for par, sbit, bsz, mode, com, id, reg in itertools.product(
                                                parities,
                                                stopbits_,
                                                bytesizes,
                                                modes,
                                                comrange,
                                                device_ids,
                                                registers):
        if id not in device_ids_to_go:
            continue
        if mode=="rtu" and bsz==7:
            continue
        comname = f"COM{com}"

        try:
            if search_count > 40:
                if search_count%10==0:
                    print(".", end='', flush=True)
                if (search_count-40)%500==0:
                    print()
            search_count+=1
            dev = minimalmodbus.Instrument(comname, id, mode=mode)
            dev.serial.parity = par
            dev.serial.stopbits = sbit
            dev.serial.bytesizes = bsz
            dev.serial.baudrate = 9600
        except serial.serialutil.SerialException:
            continue

        cont = False
        for f in (dev.read_bits, dev.read_string):
            try:
                f(reg, 1)
                if search_count>40:
                    print()
                print(f"Found dev={id} on {comname}")
                search_count = 0
                instances[id] = (comname, dev)
                device_ids_to_go.remove(id)
                break

            except (minimalmodbus.NoResponseError, minimalmodbus.InvalidResponseError, minimalmodbus.IllegalRequestError):
                pass

    not_found = set(device_ids).difference(instances)
    if not_found:
        print()
        ln = '\n'
        errmsg = f"Could not connect to Devices: {list(not_found)}"
        if instances:
            errmsg = errmsg + f", but did find: \n{ln.join([str(i) for i in instances.values()])}"
        errmsg = errmsg + "."

        raise ConnectionError(errmsg)

    return instances


if __name__ == "__main__":
    instruments = get_instruments([1,2,3,4,6])

