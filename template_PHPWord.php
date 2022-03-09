<?php
    //Verificando quais dados estão chegando do banco para utilização correta da nomeclatura
    var_dump('dados recebido do banco',$dados);

    require_once './vendor/autoload.php';
    use PhpOffice\PhpWord\Element\Section;
    use PhpOffice\PhpWord\Shared\Converter;

    //Nome que o arquivo receberá quando for baixado
    $filename = $dados->codigo ."_".$dados->empresa_nome.".docx"; 

    //Insere a data do download no arquivo
    $date= new \DateTime();

    // Cria um novo documento para inserção dos dados...
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
    
    //Padroniza fonte 
        $phpWord->setDefaultFontName('Times New Roman');
        $phpWord->setDefaultFontSize(10);
    
    
/* Note: Todo elemento deve estar dentro de uma seção */

    // Novo documento contendo Cabeçalho Corpo e Rodape
     $section = $phpWord->addSection(
         array(
             
             'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER,//centraliza o conteudo
         )

     );
     
    // Estilo do titulos
     $TitlePrimary = array('size' => 16, 'bold' => true);
     $TitleSecundary = array('size' => 12, 'bold' => true);
     
    //Cabeçalho
        $header = $section->addHeader();
        printSeparator($section);
        $source = 'assets/img/cabecalho.png'; //imagen do cabeçalho 
        $fileContent = file_get_contents($source);
        $header->addText('texto para cabeçalho caso prefira',$date);
        $header->addImage(
           $fileContent, 
           array(
             'height' => 35, 
             'width' => 455,
             'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::START
           )
        );
  
    //Rodapés
        $footer = $section->addFooter();
        $source = 'assets/img/rodape.png';
        $fileContent = file_get_contents($source);
        $footer->addPreserveText(' {PAGE} / {NUMPAGES}.'); // contador de paginas
        $footer->addImage(
           $fileContent, 
           array(
             'height' => 45, 
             'width' => 455,
             'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::START //Justifica a direita
           )
       );


    //Estilo da tabela
        $fancyTableStyle = array(
            'borderSize' => 6, 
            //'borderColor' => '006699', 
            'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::START, //Justifica a direita
        );
        $widthCell = 10000;
        $fancyTableFontStyle = array('size' => 10,'bold' => true);
        $fancyTableCellStyle = array('valign' => 'center');
        $fancyTableCellBtlrStyle = array('valign' => 'center', 'textDirection' => \PhpOffice\PhpWord\Style\Cell::TEXT_DIR_LRTB); //Justifica conteudo a direita da celula

    //Titulo do Relatório
        $title = "Meu relatório";

    // Chama o Titulo
        $section->addText(
            $title,
            $TitlePrimary,
            array(
                
                'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER,//centraliza o conteudo
            )
        );

    //Adiciona Subtitulo com nome da Empresa
        $section->addText(
            $checklist->empresa_nome,
            $TitlePrimary,
            array(
                
                'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER,//centraliza o conteudo
            )
        );

    //Cria a tabela Dados Serviço 01
        $section->addTextBreak(1);
        $section->addText('Dados Serviço 01', $TitleSecundary);
        
        $fancyTableStyleName = 'Dados Serviço 01';
        $phpWord->addTableStyle($fancyTableStyleName, $fancyTableStyle,  /*$fancyTableFirstRowStyle */);
        $table = $section->addTable($fancyTableStyleName);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Manutenção:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->codigo);
            $table->addCell($widthCell)->addText('Código serviço01:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->codigo_serviço01);
            $table->addRow(90);
            $table->addCell($widthCell)->addText('Data Agendada:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->data_agendada);
            $table->addCell($widthCell)->addText('Tipo de manuteção:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->tipo_manutencao);
            $table->addRow(90);
            $table->addCell($widthCell)->addText('Código serviço02:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->codigo_serviço02);
            $table->addCell($widthCell)->addText('Data visita:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->data_visita);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Município:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->cidade);
            $table->addCell($widthCell)->addText('Bairro:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->bairro);
            $table->addRow(90);
            $table->addCell($widthCell)->addText('Técnico:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->tecnico);

        // Adiciona as imagens 
            foreach ($fotosManutencao as $foto) {
                $section->addText(" $foto->nome");
                $section->addImage(
                    'assets/$foto->nome'.'.jpg',
                    array(
                        'positioning'        => 'relative',
                        'marginTop'          => -1,
                        'marginLeft'         => 1,
                        'width'              => 80,
                        'height'             => 80,
                        'foto'              => $foto,
                        'fotoDistanceRight'  => Converter::cmToPoint(1),
                        'wrapDistanceBottom' => Converter::cmToPoint(1),
                    )
                );
                $section->addText($text);
                printSeparator($section);
            }

    //Cria a tabela Situação na Local
        $section->addTextBreak(1);
        $section->addText('Situação na maquina', $TitleSecundary);
        
        $fancyTableStyleName = 'Situação na local';
        $phpWord->addTableStyle($fancyTableStyleName, $fancyTableStyle);
        $table = $section->addTable($fancyTableStyleName);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Hora inicial:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->hora_inicial);
            $table->addCell($widthCell)->addText('Hora final:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->hora_final);
            $table->addRow(90);
            $table->addCell($widthCell)->addText('Nível inicial da régua:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->nivel_regua_inicial);
            $table->addCell($widthCell)->addText('Nível final da régua:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->nivel_regua_final);
            $table->addRow(90);
            $table->addCell($widthCell)->addText('Réguas:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->reguas);
            $table->addCell($widthCell)->addText('Acesso:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->acesso);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Pluviômetro:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->pluviometro);
            $table->addCell($widthCell)->addText('RNs:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->rns);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Seção de medição:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->secao_medicao);
            $table->addCell($widthCell)->addText('Maneta:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->maneta);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Limpeza:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->limpeza);
            $table->addCell($widthCell)->addText('Estado geral:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->estado_geral);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('PI - PF:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->pi . '-' . $dados->pf);
            $table->addCell($widthCell)->addText('Instalação do sensor de óleo:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->sensor_oleo);
    
    
    
    //Cria a tabela Dados Dados Técnicos Hidrológicos
        $section->addTextBreak(1);
        $section->addText('Dados Técnicos Externos', $TitleSecundary);
        
        $fancyTableStyleName = 'Dados Técnicos Externos';
        $phpWord->addTableStyle($fancyTableStyleName, $fancyTableStyle);
        $table = $section->addTable($fancyTableStyleName);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Pintura:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->pintura);
            $table->addCell($widthCell)->addText('Limpeza geral:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($addos->limpeza_geral);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Nivelamento:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->nivelamento);
            $table->addCell($widthCell)->addText('Topobatimetria:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->topobatimetria);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('Vazamentos Líquidos:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->vazamentos_liquidos);
            $table->addCell($widthCell)->addText('Amostragem de sedimento (rezidos Sólido):', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->rezidos_solido);

            $table->addRow(90);
            $table->addCell($widthCell)->addText('RNs:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($checklist->rn);
            $table->addCell($widthCell)->addText('tipo graxa:', $fancyTableFontStyle);
            $table->addCell($widthCell)->addText($dados->tipo_graxa);
            
   
        
    //Cria a tabela Observações
        
        $section->addTextBreak(1);
        $section->addText('Observações', $TitleSecundary);
        
        $fancyTableStyleName = 'Observações';
        $phpWord->addTableStyle($fancyTableStyleName, $fancyTableStyle);
        $table = $section->addTable($fancyTableStyleName);

            $table->addRow(400);
            $table->addCell(20000)->addText($checklist->observacao);

    

    // Função para separar e posicionar as imagens
    function printSeparator(Section $section)
    {
        $section->addTextBreak(2);
        $lineStyle = array('weight' => 0.2, 'width' => 150, 'height' => 0, 'align' => 'center');
        $section->addLine($lineStyle);
        $section->addTextBreak(2);
    }
    
    
    /*  // Salva documento em HTML para visualização no console(F12) do browser...
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
    $objWriter->save('simples_conferencia.html'); */

    

    // Define o tipo de arquivo de saída
    header('Content-type: application/vnd.ms-word');

    // Insere o nome no arquivo
    header('Content-Disposition: attachment; filename="'.$filename.'"');

    // Disponibiliza o arquivo no browser para download
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save('php://output');