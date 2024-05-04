<?php

namespace Chuva\Php\WebScrapping;

use Chuva\Php\WebScrapping\Entity\Paper;
use Chuva\Php\WebScrapping\Entity\Person;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;


libxml_use_internal_errors(true);

/**
 * Does the scrapping of a webpage.
 */
class Scrapper {

  /**
   * Loads paper information from the HTML and returns the array with the data.
   */
    public function scrap(\DOMDocument $dom): array {
        #Caminho para o arquivo Excel.
        $filePath = 'C:\xampp\htdocs\exercicios-2024-master\php\assets\model.xlsx';

        $scrapData = $this -> scrapHTML($dom);
        $this -> Write($scrapData);

    return [];
    }

    public function scrapHTML(\DOMDocument $dom): array {
        #Coletando todo o conteudo em html.
        $content = file_get_contents("http://localhost/exercicios-2024-master/php/assets/origin.html");
        #Criando DOM e carregando o html nele.
        $document = new $dom();
        $document->loadHTML($content);
        #Utilizando o DOMXPATH para buscar elementos no html.
        $xPath = new \DOMXPath($document);

        #Selecionando a quantidade de links a serem lidos.
        $links = $xPath -> query('.//a[@class="paper-card p-lg bd-gradient-left"]');

        #Criando local de armazenamento para os dados.
        $scrapData = [];

        foreach($links as $link){
            #Definindo elementos para a busca (css convertido para xpath).

            #Selecionando apenas o primeiro elemento encontrado, usamos [0] para acessá-lo diretamente.
            $domNodeListAuthorPos = $xPath -> query('.//div[@class="authors"]', $link)[0];
            #Procurando dentro dele, todos os elementos <span> que contêm os nomes dos autores.
            $domNodeListAuthor = $xPath -> query('.//span', $domNodeListAuthorPos);
            $domNodeListTitle = $xPath -> query('.//h4[@class="my-xs paper-title"]', $link);
            $domNodeListID = $xPath -> query('.//div[@class="volume-info"]', $link);
            $domNodeListType = $xPath -> query('.//div[@class="tags mr-sm"]', $link);
            $domNodeListInstituition = $xPath -> query('.//a[@class="paper-card p-lg bd-gradient-left"]//span/@title');

            #Relacionando autores com as instituições.
            $authorArray = [];
            foreach ($domNodeListAuthor as $index => $authorNode) {
                $name = $authorNode->nodeValue;
                $instituition = $domNodeListInstituition[$index]->nodeValue;

                $person = new Person($name, $instituition);
                array_push($authorArray, $person);

            }

            #Criando array que deverá conter os dados das papers.
            $paper = new Paper($domNodeListID[0]->textContent, $domNodeListTitle[0]->textContent,$domNodeListType[0]->textContent, $authorArray);
            array_push($scrapData, $paper);
        }

        return $scrapData;
    }

    public function Write($scrapData): array {
        #Caminho para o arquivo Excel.
        $filePath = 'C:\xampp\htdocs\exercicios-2024-master\php\assets\model.xlsx';

        #Abrindo o arquivo Excel para escrita.
        $writer = WriterEntityFactory::createWriterFromFile($filePath);
        $writer->openToFile($filePath);

        #Criando o header.
        $headerRow = WriterEntityFactory::createRowFromArray(['ID', 'Title', 'Type', 'Author 1', 'Author 1 Institution', 'Author 2', 'Author 2 Institution', 'Author 3', 'Author 3 Institution', 'Author 4', 'Author 4 Institution', 'Author 5', 'Author 5 Institution', 'Author 6', 'Author 6 Institution', 'Author 7', 'Author 7 Institution', 'Author 8', 'Author 8 Institution', 'Author 9', 'Author 9 Institution']);
        $writer->addRow($headerRow);


        #Adicionando informações dos autores de cada paper ao array.
        foreach ($scrapData as $rowData) {
        $rowArray = [
            $rowData->id,
            $rowData->title,
            $rowData->type
        ];
        foreach ($rowData->authors as $author) {
            $rowArray[] = $author->name;
            $rowArray[] = $author->institution;
        }

        #Criando uma nova linha para o arquivo Excel.
        $row = WriterEntityFactory::createRowFromArray($rowArray);

        #Escrevendo o conteúdo já formatado.
        $writer->addRow($row);
        }

        #Fechar o arquivo Excel
        $writer->close();

        return [];
        
    }
}

echo "Alterado!";
