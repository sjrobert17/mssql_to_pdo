<?php
	if (!defined("MSSQL_ASSOC")) { define("MSSQL_ASSOC", "MSSQL_ASSOC"); }
	if (!defined("MSSQL_NUM")) { define("MSSQL_NUM", "MSSQL_NUM"); }
	if (!defined("MSSQL_BOTH")) { define("MSSQL_BOTH", "MSSQL_BOTH"); }
	
	function mssql_connect ($server, $dbuser, $dbpassword, $port = "1433") {
		$connection = new mssql_to_pdo_connection();
		mssql_to_pdo_connection_manager::save_connection($connection);
		return $connection->mssql_connect($server, $dbuser, $dbpassword, $port);
	}
	function mssql_select_db ($db) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_select_db($db);
	}
	function mssql_query ($query, $pdo_object = NULL) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_query($query);
	}
	function mssql_get_last_message () {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_get_last_message();
	}
	function mssql_fetch_array ($statement = NULL, $fetchstyle = MSSQL_ASSOC, $fetch_argument = NULL) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_fetch_array ($statement, $fetchstyle, $fetch_argument);
	}
	function mssql_result ($statement = NULL, $row = 0, $field = 0) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_result($statement, $row, $field);
	}
	function mssql_rows_affected ($statement = NULL) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_rows_affected($statement);
	}
	function mssql_num_fields ($result) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_num_fields($result);
	}
	function mssql_field_name ($result, $offset = 0) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_field_name($result, $offset);
	}
	function mssql_field_type ($result, $offset = -1) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_field_type($result, $offset);
	}
	function mssql_field_length ($result, $offset = -1) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_field_length($result, $offset);
	}
	function mssql_free_result ($result = NULL) {
		$connection = mssql_to_pdo_connection_manager::get_connection();
		return $connection->mssql_free_result($result);
	}
	class mssql_to_pdo_connection_manager {
		private static $connections;
		private static $connection_id = 0;
		public static function save_connection ($connection) {
			self::$connections[self::$connection_id] = $connection;
			self::$connection_id++;
		}
		public static function get_connection () {
			return self::$connections[self::$connection_id-1];
		}
	}
	class mssql_to_pdo_connection {
		public $connection_id;
		private $dsn;
		private $driver;
		private $dbname;
		private $server;
		private $user;
		private $password;
		private $pdo_object;
		
		private $last_query;
		private $last_message;
		private $query;
		private $results;
		private $result_index = 0;
		
		public function __construct () {
			static $connection_id = 0;
			$this->connection_id = $connection_id++;
		}
		public function mssql_connect ($server, $dbuser, $dbpassword, $port = "1433") {
			$this->driver = "sqlsrv";
			$this->server = $server;
			$this->user = $dbuser;
			$this->password = $dbpassword;
			$this->port = $port;
			$this->dsn = $this->driver . ":server=" . $this->server . "," . $this->port . ";ConnectionPooling=0";
			$this->pdo_object = new PDO($this->dsn, $this->user, $this->password);
			return $this->pdo_object;
		}
		public function mssql_select_db ($db) {
			$this->dbname = $db;
			$this->pdo_object->exec("USE " . $this->dbname);
			return $this->pdo_object;
		}
		public function mssql_query ($query, $pdo_object = NULL) {
			if (isset($this->query) && !empty($this->query)) {
				$this->last_query = $this->query;
			}
			$this->query = $query;
			if ($pdo_object === NULL) {
				$pdo_object = $this->pdo_object;
			}
			$statement = $pdo_object->prepare($this->query);
			$statement->execute();
			$result = $statement->fetchAll(PDO::FETCH_ASSOC);
			$rows_affected = $statement->rowCount();
			$num_fields = $statement->columnCount();
			$fields_name = ($num_fields > 0 ? array_keys($result[0]) : array());
			$fields_type = array();
			$fields_length = array();
			for ($i=0; $i < $num_fields; $i++) {
				$meta = $statement->getColumnMeta($i);
				$fields_type[$i] = $meta["sqlsrv:decl_type"];
				$fields_length[$i] = $meta["len"];
			}
			$this->results[$this->result_index] = new mssql_to_pdo_result($query, $statement, $result, $rows_affected, $num_fields, $fields_name, $fields_type, $fields_length);
			$this->result_index++;
			$this->last_message = $this->pdo_object->errorInfo();
			return $statement;
		}
		public function mssql_get_last_message () {
			return (isset($this->last_message) && !empty($this->last_message) ? $this->last_message : FALSE);
		}
		public function mssql_fetch_array ($statement = NULL, $fetchstyle = PDO::MSSQL_ASSOC, $fetch_argument = NULL) {
			return $this->get_result_set_by_statement($statement)->mssql_fetch_array($fetchstyle, $fetch_argument);
		}
		public function mssql_result ($statement = NULL, $row = 0, $field = 0) {
			return $this->get_result_set_by_statement($statement)->mssql_result($row, $field);
		}
		public function mssql_num_fields ($result) {
			return $this->get_result_set_by_result($result)->mssql_num_fields();
		}
		public function mssql_rows_affected ($statement = NULL) {
			return $this->get_result_set_by_statement($statement)->mssql_rows_affected();
		}
		public function mssql_field_name ($result, $offset = 0) {
			return $this->get_result_set_by_result($result)->mssql_field_name($offset);
		}
		public function mssql_field_type ($result, $offset = -1) {
			return $this->get_result_set_by_result($result)->mssql_field_type($offset);
		}
		public function mssql_field_length ($result, $offset = -1) {
			return $this->get_result_set_by_result($result)->mssql_field_length($offset);
		}
		public function mssql_free_result ($result = NULL) {
			if ($result === NULL) {
				unset($this->results[$this->result_index]);
				$this->result_index = $this->result_index - 1;
			}
			else {
				$count = count($this->results);
				for ($i = 0;$i < $count;$i++) {
					if ($this->results[$i]->result === $result) {
						unset($this->results[$i]);
						$this->result_index = $this->result_index - 1;
						$this->results = array_values($this->results);
					}
				}
			}
		}
		private function get_result_set_by_statement ($statement) {
			$count = count($this->results);
			for ($i = 0;$i < $count;$i++) {
				if ($this->results[$i]->statement === $statement) {
					return $this->results[$i];
				}
			}
			return ($this->results > 0 ? $this->results[$this->result_index-1] : FALSE);
		}
		private function get_result_set_by_result ($result) {
			$count = count($this->results);
			for ($i = 0;$i < $count;$i++) {
				if ($this->results[$i]->result === $result) {
					return $this->results[$i];
				}
			}
			return ($this->results > 0 ? $this->results[$this->result_index-1] : FALSE);
		}
	}
	class mssql_to_pdo_result {
		public $result_id;
		public $query;
		public $statement;
		public $result;
		public $rows_affected;
		public $num_fields;
		public $fields_name;
		public $fields_type;
		public $fields_length;
		public function __construct ($query, $statement, $result, $rows_affected, $num_fields, $fields_name, $fields_type, $fields_length) {
			static $result_id = 0;
			$this->result_id = $result_id++;
			$this->query = $query;
			$this->statement = $statement;
			$this->result = $result;
			$this->rows_affected = $rows_affected;
			$this->num_fields = $num_fields;
			$this->fields_name = $fields_name;
			$this->fields_type = $fields_type;
			$this->fields_length = $fields_length;
		}
		public function mssql_fetch_array ($fetchstyle = MSSQL_ASSOC, $fetch_argument = NULL) {
			if ($fetch_argument !== NULL) {
				if ($fetch_argument == NULL || !(is_integer($fetch_argument) || is_string($fetch_argument))) {
					$fetch_argument = 0;
				}
				$assoc_array = $this->result;
				$indexed_array = array_map(array_values, $assoc_array);
				$target_array = (is_integer($fetch_argument) ? $indexed_array : $assoc_array);
				$return = array();
				foreach ($target_array as $key => $value) {
					$return[] = $value[$fetch_argument];
				}
				return $return;
			}
			else if ($fetchstyle == MSSQL_ASSOC) {
				return $this->result;
			}
			else if ($fetchstyle == MSSQL_NUM) {
				$assoc_array = $this->result;
				$indexed_array = array_map(array_values, $assoc_array);
				return $indexed_array;
			}
			else if ($fetchstyle == MSSQL_BOTH) {
				$assoc_array = $this->result;
				$indexed_array = array_map(array_values, $assoc_array);
				$both_array = array();
				for ($i = 0;$i < count($assoc_array);$i++) {
					$both_array[$i] = array_merge($assoc_array[$i], $indexed_array[$i]);
				}
				return $both_array;
			}
			else {
				return $this->result;
			}
		}
		public function mssql_result ($row = 0, $field = 0) {
			if (is_integer($field)) {
				$array = $this->mssql_fetch_array(MSSQL_NUM, $field);
				$result = $array[$row];
			}
			else {
				$assoc_array = $this->mssql_fetch_array(MSSQL_ASSOC);
				$result = $assoc_array[$row][$field];
			}
			return $result;
		}
		public function mssql_num_fields () {
			return $this->num_fields;
		}
		public function mssql_rows_affected () {
			return $this->rows_affected;
		}
		public function mssql_field_name ($offset = 0) {
			if (is_integer($offset) && $offset > -1 && $offset < $this->num_fields) {
				return $this->fields_name[$offset];
			}
			else {
				return FALSE;
			}
		}
		public function mssql_field_type ($offset = -1) {
			return $this->fields_type[$offset];
		}
		public function mssql_field_length ($offset = -1) {
			return $this->fields_length[$offset];
		}
	}
?>