<?php
if (!defined('BASEPATH')) exit('No direct script access allowed');

class Excel_data_importer {

    private $CI     = NULL;
    private $import_excel_files = array();

    public function __construct()
    {
        $this->CI = &get_instance();
    }

    public function import($import_excel_files=array()) {
        $this->import_excel_files = $import_excel_files;

        //データをすべて削除
        //$this->truncae_tables();
        //実行
        $this->load_excel_data($this->import_excel_files);
    }

    public function begin() {
        //トランザクション開始
        $this->CI->db->trans_begin();
    }

    public function end($commit_data=false) {
        if ( $commit_data ) {
            $this->CI->db->trans_complete();
        } else {
            //トランザクション終了（ロールバック）
            $this->CI->db->trans_rollback();
        }
    }

    /**
     * 指定したExcelファイルを読み込んでDBに投入
     */
    public function load_excel_data($excel_files) {
        if ( $excel_files === null || count($excel_files) < 1 ) return;
        foreach ($excel_files as $excel_file_path) {
            if ( file_exists($excel_file_path) ) {
                $this->insert_from_excel($excel_file_path);
            }
        }
    }


    public function truncae_tables() {
        $tables = $this->CI->db->list_tables();
        foreach ($tables as $table_name) {
            if ( strpos($table_name, 'view_') === 0 ) continue;
            //全削除
            $this->CI->db->empty_table($table_name);
        }
    }

    private function insert_from_excel($excel_file_path) {
        $file_ext = pathinfo($excel_file_path, PATHINFO_EXTENSION);

        if ( $file_ext === 'xlsx' ) {
            $reader = PHPExcel_IOFactory::createReader('Excel2007');
        } else {
            $reader = PHPExcel_IOFactory::createReader('Excel5');
        }
        $reader->setReadDataOnly(true);

        //Excel読み込み
        $excel_file = $reader->load($excel_file_path);
        //シート取得
        $sheet_names = $excel_file->getSheetNames();
        foreach ($sheet_names as $sheet_name) {
            $sheet = $excel_file->getSheetByName($sheet_name);

            $this->insert_from_sheet($sheet);
        }
    }

    private function get_table_schema($table_name) {
        $result = array();

        $table_schema = $this->CI->db->field_data($table_name);
        foreach ($table_schema as $schema) {
            $result[$schema->name] = $schema;
        }
        return $result;
    }

    private function insert_from_sheet($excel_sheet) {
        $table_name = $excel_sheet->getTitle();
        $table_schema = $this->get_table_schema($table_name);

        $table_data = $this->get_excel_data($excel_sheet, $table_schema);
        //insert batch
        $this->CI->db->insert_batch($table_name, $table_data);
    }

    private function get_excel_data($excel_sheet, $table_schema) {
        $excel_headers = null;
        $excel_data = array();
        foreach ( $excel_sheet->getRowIterator() as $row ) {
            if ( $row->getRowIndex() == 1 ) {
                //ヘッダ取得
                $excel_headers = $this->get_values($row);
                continue;
            }

            $excel_data[] = $this->get_values($row, $excel_headers, $table_schema);
        }
        //
        return $excel_data;
    }


    private function convert_data($excel_data, $table_schema) {
        $data = array();
        foreach ($excel_data as $key => $value) {
            if ( isset($table_schema[$key]) ) {
                $schema = $table_schema[$key];
                $type = $schema->type;
                if ( $type ==="date" || $type ==="datetime" ) {
                    if ( !is_null($value) ) {
                        $value = is_numeric($value) ? date('Y-m-d', $value) : date('Y-m-d', strtotime($value));
                    }
                }
            }
            $data[$key] = $value;
        }
        return $data;
    }

    private function get_values($row, $fields=null, $table_schema=null) {
        $data = array();
        if ( $fields ) {
            foreach ($fields as $key) {
                $data[$key] = NULL;
            }
        }
        //列でループ
        foreach ( $row->getCellIterator() as $key => $cell ) {
            if ( ! is_null($cell) ) {

                $type = "text";

                //$fieldsが指定されている場合は、フィールド名の連想配列とする
                if ( $fields != null && isset($fields[$key]) ) $key = $fields[$key];
                if ( $table_schema != null && isset($table_schema[$key]) ) $type = $table_schema[$key]->type;
                //
                if ( isset($this->schema_type_map[$type]) ) $type = $this->schema_type_map[$type];

                $data[$key] = $this->get_value($cell, $type);
            }
        }
        return $data;
    }

    private function get_value($cell, $type) {
        //var_dump($type);
        $value = $cell->getValue();
        if ( is_null( $value ) ) return null;
        if( $type == "date" ) {
            $value = PHPExcel_Style_NumberFormat::toFormattedString($value, 'yyyy-mm-dd');
        } else if ( $type == "timestamp" ) {
            $value = PHPExcel_Style_NumberFormat::toFormattedString($value, 'yyyy-mm-dd h:mm:ss');
        } else if ( $type == "text" ) {
            $value = '' . $value;
        }
        return $value;
    }

    private $schema_type_map = array(
        "date" => "date",
        "timestamp without time zone" => "timestamp",
        "text" => "text",
        "character varying" => "text",
    );
}
