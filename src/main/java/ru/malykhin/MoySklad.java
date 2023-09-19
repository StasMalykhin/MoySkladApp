package ru.malykhin;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.web.client.RestTemplate;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

public class MoySklad {
    private static RestTemplate restTemplate = new RestTemplate();
    private static ObjectMapper objectMapper = new ObjectMapper();

    public static void main(String[] args) throws IOException {
        if (args.length == 1) {
            String input = args[0];

            String configFilePath = "application.properties";
            Properties properties = new Properties();
            InputStream configFile = MoySklad.class.getClassLoader().getResourceAsStream(configFilePath);
            properties.load(configFile);

            String accessToken = properties.getProperty("moysklad.accessToken");

            File inputFile = Path.of(input).toFile();
            int currentRow = 0;
            int nextRow = 1;

            HttpEntity<HttpHeaders> request = getHeaders(accessToken);

            //Формируем список складов, которые сейчас определены в системе
            Map<String, String> stocks = createListStocks(request);

            while (ConnectExcel.readFromExcel(inputFile, currentRow) != null) {

                //Вытаскиваем из файла Excel название товара
                String nameProduct = ConnectExcel.readFromExcel(inputFile, currentRow);

                //По названию товара находим его id через MoySklad API
                String productId = findProductId(nameProduct, request);

                //Для хранения списка остатков с разбивкой по складам
                Map<String, String> quantityInStocks = new HashMap<>();

                if (productId != null) {

                    //Формируем для данного товара список остатков
                    quantityInStocks = createListForQuantityInStocks(productId,
                            request, stocks);
                }
                //Записываем в файл Excel наименование товара и остатки на складах
                ConnectExcel.writeIntoExcel(inputFile, nextRow, nameProduct, stocks, quantityInStocks);
                currentRow++;
                nextRow++;
            }
        } else {
            System.out.println("Incorrect number of arguments passed to program");
        }
    }

    private static HttpEntity<HttpHeaders> getHeaders(String accessToken) {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_JSON);
        headers.add("Authorization", "Basic " + accessToken);
        return new HttpEntity<>(headers);
    }

    private static Map<String, String> createListStocks(HttpEntity<HttpHeaders> request) throws JsonProcessingException {
        String url1 = "https://online.moysklad.ru/api/remap/1.2/entity/store";

        String response1 = restTemplate.exchange(url1, HttpMethod.GET, request, String.class).getBody();
        JsonNode obj = objectMapper.readTree(response1);

        Map<String, String> stocks = new HashMap<>();

        for (int i = 0; i < obj.get("rows").size(); i++) {
            String stockId = obj.get("rows").get(i).get("id").asText();
            String nameStock = obj.get("rows").get(i).get("name").asText();

            stocks.put(nameStock, stockId);
        }
        return stocks;
    }

    private static String findProductId(String nameProduct, HttpEntity<HttpHeaders> request) throws JsonProcessingException {
        String url2 = "https://online.moysklad.ru/api/remap/1.2/entity/" +
                "product?filter=code=" + nameProduct;

        String response2 = restTemplate.exchange(url2, HttpMethod.GET,
                request, String.class).getBody();
        JsonNode obj = objectMapper.readTree(response2);

        if (!obj.get("rows").isEmpty()) {
            return obj.get("rows").get(0).get("id").asText();
        }
        return null;
    }

    private static Map<String, String> createListForQuantityInStocks(String productId, HttpEntity<HttpHeaders> request, Map<String, String> stocks) throws JsonProcessingException {

        String url3 = "https://online.moysklad.ru/api/remap/1.2/report/" +
                "stock/bystore/current?filter=assortmentId=" + productId;

        String response3 = restTemplate.exchange(url3, HttpMethod.GET,
                request, String.class).getBody();
        JsonNode obj = objectMapper.readTree(response3);

        Map<String, String> quantityInStocks = new HashMap<>();

        for (int i = 0; i < obj.size(); i++) {
            String selectedStockId = obj.get(i).get("storeId").asText();

            for (String nameStock : stocks.keySet()) {
                String stockId = stocks.get(nameStock);

                if (selectedStockId.equals(stockId)) {
                    String stock = obj.get(i).get("stock").toString();
                    quantityInStocks.put(nameStock, stock);
                }
            }
        }
        return quantityInStocks;
    }

}

