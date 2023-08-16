<?php

namespace App\Http\Controllers;


use Illuminate\Http\Request;
use ZipArchive;

class Excel extends Controller
{
    public function upload(Request $request)
    {
        header('Content-Type: text/html; charset=utf-8');
        $file = $request->file('sheet');
        if (!empty($request->input("predio"))) {
            $predio = $request->input("predio");
        } else {
            //ERRO
        }
        if (!empty($request->input("local"))) {
            $local = $request->input("local");
        } else {
            //ERRO
        }
        if ($file) {
            $filename = $file->getClientOriginalName();
            $extension = $file->getClientOriginalExtension();
            $tempPath = $file->getRealPath();
            $fileSize = $file->getSize();
            $location = 'uploads';
            $file->move($location, $filename);
            $filepath = public_path($location . "/" . $filename);
            $file = fopen($filepath, "r");
            $csv = [];
            $data = [];
            if (($open = fopen($filepath, "r")) !== FALSE) {

                while ($line = fgets($open)) {
                    $data[] = $line;
                }
                fclose($open);
            }
            $con = 0;
            foreach ($data as $dd) {
                $csv[$con] = explode(';', $dd);
                $con++;
            }

            $importITE = [];
            $importITE[0] = ['C'];
            $importITE[1] = ['command', 'subGroup', 'itemCategory', 'description', 'alternativeIdentifier', 'active', 'CF_equipamento', 'CF_LOCAL', 'CF_LOCAL_DO_ITEM', 'CF_Periodo', 'CF_TAG', 'CF_AREA', 'CF_tipo_eq'];
            $importISA = [];
            $importISA[0] = ['C'];
            $importISA[1] = ['command', 'active', 'order', 'section', 'item'];
            $con = 2;
            unset($csv[0]);
            $csv = array_values($csv);
            foreach ($csv as $datacsv) {
                if($datacsv[0] != ""){
                    $alt = $datacsv[0];
                    $tipo = $datacsv[1];
                    $pavi = $datacsv[2];
                    $peri = $datacsv[3];
                    $predio = mb_convert_encoding($predio, 'UTF-8');
                    $alt = mb_convert_encoding($alt, 'UTF-8');
                    $tipo = mb_convert_encoding($tipo, 'UTF-8');
                    $pavi = mb_convert_encoding($pavi, 'UTF-8');
                    $peri = mb_convert_encoding($peri, 'UTF-8');
                    $con = mb_convert_encoding($con, 'UTF-8');
                    $local = mb_convert_encoding($local, 'UTF-8');
                    $importITE[$con] = Excel::fnITE($predio, $alt, $tipo, $pavi, $peri, $con, $local);
                    $importISA[$con] = Excel::fnISA($predio, $alt, $tipo, $pavi, $peri, $con, $local);
                    $con++;
                }
            }
            $f = array(Excel::fileCSV($importITE, $filename = "ITE_V2.txt"),  Excel::fileCSV($importISA, $filename = "ISA_V2.txt"));
            Excel::createZipFile($f, $fileNameDownload = "importHersil");
        }
    }
    public function fnITE($predio, $alt, $tipo, $pavi, $peri, $cont, $local)
    {
        $predio = strtoupper($predio);
        $command = "I";
        $subGroup = $tipo;
        $itemCategory = "";
        switch ($tipo) {
            case "Ancoragem":
                $itemCategory = "Civil";
                break;
            case "Automação de portões":
                $itemCategory = "automacao";
                break;
            case "Barrilete":
                $itemCategory = "Hidráulica";
                break;
            case "Bombas de água e esgoto":
                $itemCategory = "Hidráulica";;
                break;
            case "Bombas de incêndio":
                $itemCategory = "Incêndio";
                break;
            case "Caixas de Esgoto":
                $itemCategory = "Hidráulica";
                break;
            case "Casa de Máquinas":
                $itemCategory = "Elevadores";
                break;
            case "Cftv, sca, sai e bms":
                $itemCategory = "automacao";
                break;
            case "Dedetização e desratização":
                $itemCategory = "Civil";
                break;
            case "Elevadores":
                $itemCategory = "Elevadores";
                break;
            case "Esquadrias":
                $itemCategory = "Civil";
                break;
            case "Extintores":
                $itemCategory = "Incêndio";
                break;
            case "Iluminação de emergência":
                $itemCategory = "Incêndio";
                break;
            case "Impermeabilização":
                $itemCategory = "Civil";
                break;
            case "Jardim":
                $itemCategory = "Jardim";
                break;
            case "Limpeza Fachada":
                $itemCategory = "Civil";
                break;
            case "Mangueiras":
                $itemCategory = "Incêndio";
                break;
            case "Metais, acessórios e registros":
                $itemCategory = "Hidráulica";
                break;
            case "Nobreaks":
                $itemCategory = "Eletrica";
                break;
            case "Paredes externas, fachada e muros":
                $itemCategory = "Civil";
                break;
            case "Pintura":
                $itemCategory = "Civil";
                break;
            case "Piso cimentado":
                $itemCategory = "Civil";
                break;
            case "Rejuntamento e vedações":
                $itemCategory = "Civil";
                break;
            case "Revestimentos Cerâmicos, Mármores e Granitos":
                $itemCategory = "Civil";
                break;
            case "Rufos":
                $itemCategory = "Civil";
                break;
            case "Sistema de cobertura":
                $itemCategory = "Civil";
                break;
            case "SPDA":
                $itemCategory = "Civil";
                break;
            case "Vidros e seus sistemas de fixação":
                $itemCategory = "Civil";
                break;
            case "Catracas":
                $itemCategory = "automacao";
                break;
            case "Quadro de distribuição de circuitos":
                $itemCategory = "Eletrica";
                break;
            case "Tomadas, interruptores e pontos de luz":
                $itemCategory = "Eletrica";
                break;
            case "Grupo Gerador":
                $itemCategory = "Eletrica";
                break;
            case "Porta corta-fogo":
                $itemCategory = "Incêndio";
                break;
            case "Sistema de Hidrantes e SPK":
                $itemCategory = "Incêndio";
                break;
            case "Sistema de Incêndio":
                $itemCategory = "Incêndio";
                break;
            case "Pressurização de escada":
                $itemCategory = "Incêndio";
                break;
            case "Ralos, grelhas, calhas e canaletas":
                $itemCategory = "Hidráulica";
                break;
            case "Registros de água e incêndio":
                $itemCategory = "Hidráulica";
                break;
            case "Reservatórios de água potável":
                $itemCategory = "Hidráulica";
                break;
            case "Tubulações":
                $itemCategory = "Hidráulica";
                break;
            case "Válvulas redutoras de pressão":
                $itemCategory = "Hidráulica";
                break;
            case "Rondas Diárias":
                $itemCategory = "Rotinas";
                break;
            case "Sistema de irrigação":
                $itemCategory = "Jardim";
                break;
            case "Ventiladores e Exaustores":
                $itemCategory = "Ventilação";
                break;
            case "Ventokit":
                $itemCategory = "Ventilação";
                break;
        }
        $descp = $alt;
        $peri = strtoupper(trim($peri));
        if ($peri === "SEMANAL") {
            $periodicidade = "Se";
        } else {
            $periodicidade = strtoupper(substr($peri, 0, 1));
        }
        $alternativeID = $predio . substr($descp, 0, 3) . $cont . "_" . $periodicidade;
        $active = 1;
        $cfEquip = $tipo;
        $cfLocal = $pavi;
        $cfLocalItem = $local;
        $cfPeriodo = $periodicidade;
        $cfTag = $predio . substr($descp, 0, 3) . $cont . "_" . $periodicidade;
        $cfArea = $pavi;
        $cfTipoEq = $subGroup;
        $content = [$command, $subGroup, $itemCategory, $descp, $alternativeID, $active, $cfEquip, $cfLocal, $cfLocalItem, $cfPeriodo, $cfTag, $cfArea, $cfTipoEq];
        return $content;
    }
    public function fnISA($predio, $alt, $tipo, $pavi, $peri, $cont, $local)
    {
        $command = "I";
        $active = 1;
        $order = 1;
        $peri = strtoupper(trim($peri));
        if ($peri === "SEMANAL") {
            $periodicidade = "Se";
        } else {
            $periodicidade = strtoupper(substr($peri, 0, 1));
        }
        $any = $tipo . "_" . $periodicidade;
        $section = "";
        switch ($any) {
            case "Ancoragem_A":
                $section = "ANCORAGEM_VERIFICACOES_ANUAL";
                break;
            case "Automação de portões_M":
                $section = "AUTOMACAOPORTAO_VERIFICACOES_MENSAL";
                break;
            case "Automação de portões_T":
                $section = "AUTOMACAOPORTAO_VERIFICACOES_TRIMESTRAL";
                break;
            case "Barrilete_M":
                $section = "VERIFICACAO_BARRILETE_MENSAL";
                break;
            case "Bombas de água e esgoto_M":
                $section = "VERIFICACAO_BOMBAAGUA_MENSAL";
                break;
            case "Bombas de água e esgoto_S":
                $section = "VERIFICACAO_BOMBAAGUA_SEMESTRAL";
                break;
            case "Bombas de incêndio_M":
                $section = "VERIFICACAO_BOMBAINCENDIO_MENSAL";
                break;
            case "Caixas de Esgoto_S":
                $section = "VERIFICACAO_CXESGOTO_SEMESTRAL";
                break;
            case "Casa de Máquinas_M":
                $section = "VERIFICACAO_CASAMAQUINA_MENSAL";
                break;
            case "Catracas_M":
                $section = "VERIFICACAO_CATRACA_MENSAL";
                break;
            case "Catracas_S":
                $section = "VERIFICACAO_CATRACA_SEMESTRAL";
                break;
            case "Cftv, sca, sai e bms_M":
                $section = "VERIFICACAO_DADOS_MENSAL";
                break;
            case "Cftv, sca, sai e bms_S":
                $section = "VERIFICACAO_DADOS_SEMESTRAL";
                break;
            case "Dedetização e desratização_A":
                $section = "VERIFICACAO_DEDETIZACAO_ANUAL";
                break;
            case "Dedetização e desratização_T":
                $section = "VERIFICACAO_DEDETIZACAO_TRIMESTRAL";
                break;
            case "Elevadores_M":
                $section = "VERIFICACAO_ELEVADORES_MENSAL";
                break;
            case "Elevadores_S":
                $section = "VERIFICACAO_ELEVADORES_SEMESTRAL";
                break;
            case "Elevadores_Se":
                $section = "VERIFICACAO_ELEVADORES_SEMANAL";
                break;
            case "Esquadrias_A":
                $section = "VERIFICACAO_ESQUADRIAS_ANUAL";
                break;
            case "Esquadrias_T":
                $section = "VERIFICACAO_ESQUADRIAS_TRIMESTRAL";
                break;
            case "Extintores_A":
                $section = "VERIFICACAO_EXTINTORES_ANUAL";
                break;
            case "Extintores_M":
                $section = "VERIFICACAO_EXTINTORES_MENSAL";
                break;
            case "Grupo Gerador_A":
                $section = "VERIFICACAO_GERADOR_ANUAL";
                break;
            case "Grupo Gerador_M":
                $section = "VERIFICACAO_GERADOR_MENSAL";
                break;
            case "Grupo Gerador_Q":
                $section = "VERIFICACAO_GERADOR_QUINZENAL";
                break;
            case "Grupo Gerador_S":
                $section = "VERIFICACAO_GERADOR_SEMESTRAL";
                break;
            case "Iluminação de emergência_B":
                $section = "VERIFICACAO_ILUMINACAOEMERGENCIA_BIMESTRAL";
                break;
            case "Iluminação de emergência_M":
                $section = "VERIFICACAO_ILUMINACAOEMERGENCIA_MENSAL";
                break;
            case "Impermeabilização_A":
                $section = "VERIFICACAO_IMPERMEABILIZACAO_ANUAL";
                break;
            case "Jardim_M":
                $section = "VERIFICACAO_JARDIM_MENSAL";
                break;
            case "Limpeza Fachada_A":
                $section = "LIMPEZAFACHADA_VERIFICACOES_ANUAL";
                break;
            case "Limpeza Fachada_S":
                $section = "LIMPEZAFACHADA_VERIFICACOES_SEMESTRAL";
                break;
            case "Mangueiras_A":
                $section = "VERIFICACAO_MANGUEIRA_ANUAL";
                break;
            case "Mangueiras_S":
                $section = "VERIFICACAO_MANGUEIRA_SEMESTRAL";
                break;
            case "Metais, acessórios e registros_A":
                $section = "VERIFICACAO_METAL_ANUAL";
                break;
            case "Metais, acessórios e registros_S":
                $section = "VERIFICACAO_METAL_SEMESTRAL";
                break;
            case "Nobreaks_A":
                $section = "NOBREAKS_VERIFICACOES_ANUAL";
                break;
            case "Paredes externas, fachada e muros_A":
                $section = "VERIFICACAO_PAREDES_ANUAL";
                break;
            case "Pintura_A":
                $section = "VERIFICACAO_PINTURA_ANUAL";
                break;
            case "Piso cimentado_S":
                $section = "VERIFICACAO_PISO_SEMESTRAL";
                break;
            case "Porta corta-fogo_T":
                $section = "VERIFICACAO_PORTACORTAFOGO_TRIMESTRAL";
                break;
            case "Pressurização de escada_M":
                $section = "VERIFICACAO_PRESSURIZACAO_MENSAL";
                break;
            case "Quadro de distribuição de circuitos_A":
                $section = "VERIFICACAO_QUADROS_ANUAL";
                break;
            case "Quadro de distribuição de circuitos_S":
                $section = "VERIFICACAO_QUADROS_SEMESTRAL";
                break;
            case "Quadro de distribuição de circuitos_T":
                $section = "VERIFICACAO_QUADROS_TRIMESTRAL";
                break;
            case "Ralos, grelhas, calhas e canaletas_M":
                $section = "VERIFICACAO_RALOS_MENSAL";
                break;
            case "Registros de água e incêndio_S":
                $section = "VERIFICACAO_REGISTROAGUAINCE_SEMESTRAL";
                break;
            case "Rejuntamento e vedações_A":
                $section = "VERIFICACAO_REJUNTAMENTOS_ANUAL";
                break;
            case "Reservatórios de água potável_S":
                $section = "VERIFICACAO_AGUAPOTAVEL_SEMESTRAL";
                break;
            case "Reservatórios de água potável_Se":
                $section = "VERIFICACAO_AGUAPOTAVEL_SEMANAL";
                break;
            case "Revestimentos Cerâmicos, Mármores e Granitos_A":
                $section = "VERIFICACOES_REVESTIMENTO_ANUAL";
                break;
            case "Rondas Diárias_D":
                $section = "VERIFICACOES_RONDAS_DIARIAS";
                break;
            case "Rufos_A":
                $section = "VERIFICACAO_RUFOS_ANUAL";
                break;
            case "Sistema de cobertura_S":
                $section = "VERIFICACAO_COBERTURA_SEMESTRAL";
                break;
            case "Sistema de Hidrantes e SPK_M":
                $section = "VERIFICACAO_HIDRANTE_MENSAL";
                break;
            case "Sistema de Incêndio_T":
                $section = "VERIFICACAO_SISTEMAINCENDIO_TRIMESTRAL";
                break;
            case "Sistema de irrigação_Se":
                $section = "VERIFICACAO_SISTEMAIRRIG_SEMANAL";
                break;
            case "SPDA_A":
                $section = "VERIFICACAO_SPDA_ANUAL";
                break;
            case "Tomadas, interruptores e pontos de luz_A":
                $section = "VERIFICACOES_TOMADAS_ANUAL";
                break;
            case "Tubulações_A":
                $section = "VERIFICACAO_TUBULACOES_ANUAL";
                break;
            case "Tubulações_M":
                $section = "VERIFICACAO_TUBULACOES_MENSAL";
                break;
            case "Válvulas redutoras de pressão_S":
                $section = "VERIFICACAO_VALVULAS_SEMESTRAL";
                break;
            case "Ventiladores e Exaustores_M":
                $section = "VERIFICACAO_VENTILADORES_MENSAL";
                break;
            case "Ventokit_A":
                $section = "VERIFICACAO_VENTOKIT_ANUAL";
                break;
            case "Vidros e seus sistemas de fixação_A":
                $section = "VERIFICACAO_VIDROS_ANUAL";
                break;
            case "Tubulações_M":
                $section = "VERIFICACAO_TUBULACOES_MENSAL";
                break;
            case "Válvulas redutoras de pressão_S":
                $section = "VERIFICACAO_VALVULAS_SEMESTRAL";
                break;
        }
        $descp = $alt;
        $item = $predio . substr($descp, 0, 3) . $cont . "_" . $periodicidade;
        $content = [$command, $active, $order, $section, $item];
        return $content;
    }

    public function fileCSV($array, $filename, $delimiter = ";")
    {
        $cont = 0;
        $contents = "";

        foreach ($array as $data) {
            $content = "";
            if ($cont <= 0) {
                $contents .= $data[0] . PHP_EOL;
            } else {
                foreach ($data as $line) {
                    $content .= $line . $delimiter;
                }
                $trim = rtrim($content, $delimiter);
                $contents = $contents . $trim . PHP_EOL;
            }

            $cont++;
        }
        return $contents;
    }
    function createZipFile($f = array(), $fileName)
    {
        $zip = new ZipArchive();
        $cont = 0;
        if ($zip->open("$fileName.zip", ZipArchive::CREATE) === TRUE) {
            foreach ($f as $file) {
                $zip->addFromString('import_'.$cont.'.txt', $file);
                $cont++;
            }
            $zip->addFromString('INSTRUÇÃO.txt', "Renomear os arquivos para ITE_V2.txt e ISA_V2.txt (de acordo com a segunda linha do arquivo)".PHP_EOL."ITE = command;subGroup;itemCategory;description;alternativeIdentifier;active;CF_equipamento;CF_LOCAL;CF_LOCAL_DO_ITEM;CF_Periodo;CF_TAG;CF_AREA;CF_tipo_eq".PHP_EOL."ISA = command;active;order;section;item");
            $zip->close();
        }
        header("Content-type: application/zip");
        header("Content-Disposition: attachment; filename=$fileName.zip");
        header("Content-length: " . filesize("$fileName.zip"));
        header("Pragma: no-cache");
        header("Expires: 0");
        readfile("$fileName.zip");
        unlink("$fileName.zip");
    }
}