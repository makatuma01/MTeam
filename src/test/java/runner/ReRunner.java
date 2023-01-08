package runner;


import com.google.gson.internal.bind.util.ISO8601Utils;
import io.cucumber.junit.Cucumber;
import io.cucumber.junit.CucumberOptions;
import org.junit.runner.RunWith;

@RunWith(Cucumber.class)
@CucumberOptions(

        features = "@target/rerun.txt",
        glue = "steps"
)




public class ReRunner {

}
