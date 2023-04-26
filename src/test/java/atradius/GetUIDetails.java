package atradius;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

public class GetUIDetails {

    public Map<String, String> prepareTestData(String jobName) {
        System.out.println("Input Job Name : "+jobName);
        Map<String, String> map = new LinkedHashMap<>();
        for(int i=1;i<3;i++) {
            map.put("Job Name_"+i, jobName);
            map.put("Start Date Name_"+i, jobName+" "+i+" row Start Date mine ");
            map.put("End Date Name_"+i, jobName+" "+i+" row End Date mine ");
            map.put("Record Passed_"+i, jobName+" "+i+" row Record Passed mine ");

        }
        return map;
    }
}
