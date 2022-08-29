package TellerTotaalOntkleeft;

public enum VehicleType {
    BUS(1, 6199),
    TRAM(6200, 9999),
    POLDER(10000, 1000000);

    public final int beginNumber;
    public final int endNumber;

    VehicleType(int begin, int end) {
        beginNumber = begin;
        endNumber = end;
    }
}