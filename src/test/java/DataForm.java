import java.util.HashMap;
import java.util.Map;

public class DataForm {
    private Map<String, String> map = new HashMap();

    public void add(String key, String value) {
        map.put(key, value);
    }

    public String get(String key) {
        return map.get(key);
    }

    public void print() {
        System.out.println(map);
    }

    public void replace(String key, String value) {
        map.remove(key);
        map.put(key, value);
    }
}
