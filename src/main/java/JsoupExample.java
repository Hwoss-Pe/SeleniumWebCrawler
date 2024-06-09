import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.IOException;

public class JsoupExample {
    public static void main(String[] args) {
        try {
            // 发送HTTP GET请求并获取网页内容
            Document document = Jsoup.connect("https://xz.chsi.com.cn/occupation/index.action").get();

            // 获取网页标题
            String title = document.title();
            System.out.println("网页标题：" + title);

            // 获取所有的链接
            Elements links = document.select("a[href]");
            System.out.println("链接数量：" + links.size());

            // 打印每个链接的文本和URL
            for (Element link : links) {
                String linkText = link.text();
                String linkUrl = link.attr("href");
                System.out.println("链接文本：" + linkText);
                System.out.println("链接URL：" + linkUrl);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
