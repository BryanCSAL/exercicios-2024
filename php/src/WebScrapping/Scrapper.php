<?php

namespace Chuva\Php\WebScrapping;

use Chuva\Php\WebScrapping\Entity\Paper;
use Chuva\Php\WebScrapping\Entity\Person;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Common\Entity\Row;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Common\Type;


libxml_use_internal_errors(true);

/**
 * Does the scrapping of a webpage.
 */
class Scrapper {

  /**
   * Loads paper information from the HTML and returns the array with the data.
   */
    public function scrap(\DOMDocument $dom): array {

        $scrapData = this -> scrapHTML($dom);
        $this -> Write($filePath, $scrapData);

    return [];
    }

    public function scrapHTML(\DOMDocument $dom): array {
        #Coleta todo o conteudo em html.
        $content = file_get_contents("http://localhost/exercicios-2024-master/php/assets/origin.html");
        #Criando DOM e carregando o html nele.
        $document = new DOMDocument();
        $document->loadHTML($content);
        #Utilizando o DOMXPATH para buscar elementos no html.
        $xPath = new DOMXPath($document);

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

            #Relacionando autores com as instituições
            $author_instituition = [];
            foreach ($domNodeListAuthor as $index => $authorNode) {
                $name = $authorNode->nodeValue;
                $instituition = $domNodeListInstituition[$index]->nodeValue;

                $person = new Person($author, $instituition);
                array_push($author_instituition, $person);

            }

            $paper = new Paper($domNodeListID[0]->textContent, $domNodeListTitle[0]->textContent,$domNodeListType[0]->textContent, $author_instituition);
            array_push($scrapData, $paper);
        }

        return $scrapData;
    }

    public function Write($scrapData): array {
        #Abrir o arquivo Excel para leitura
        $reader = ReaderEntityFactory::createReaderFromFile($filePath);
        $reader->open($filePath);

        #Inicializar o contador de linha
        $rowIndex = 0;

        #Abrir o arquivo Excel para escrita
        $writer = WriterEntityFactory::createWriterFromFile($filePath);
        $writer->openToFile($filePath);

        #Iterar sobre as linhas existentes do arquivo Excel
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                #Incrementar o contador de linha
                $rowIndex++;

                #Pular a primeira linha (linha do cabeçalho)
                if ($rowIndex === 1) {
                    $writer->addRow($row);
                    continue;
                }

                #Verificar se há dados correspondentes para esta linha
                if (isset($scrapData[$rowIndex])) {
                    #Substituir os dados existentes pelos novos dados
                    $row->setCells($scrapData[$rowIndex]);
                }

                #Escrever a linha atual no arquivo Excel
                $writer->addRow($row);
            }
        }

        #Escrever os dados restantes que não existem no arquivo Excel original
        for ($i = $rowIndex + 1; $i <= count($scrapData); $i++) {
            $writer->addRow(new Row($scrapData[$i]));
        }

        #Fechar o arquivo Excel
        $writer->close();
        $reader->close();

        return [];  
    }
}

$filePath = 'C:\xampp\htdocs\exercicios-2024-master\php\assets\model.xlsx';

echo "Alterado!";
