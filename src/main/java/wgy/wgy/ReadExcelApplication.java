package wgy.wgy;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@ComponentScan("wgy")
public class ReadExcelApplication {

    public static void main(String[] args) {
        SpringApplication.run(ReadExcelApplication.class, args);
    }

}
