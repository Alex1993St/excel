-- --------------------------------------------------------
-- Хост:                         127.0.0.1
-- Версия сервера:               5.7.19 - MySQL Community Server (GPL)
-- Операционная система:         Win32
-- HeidiSQL Версия:              9.4.0.5125
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!50503 SET NAMES utf8mb4 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


-- Дамп структуры базы данных excel
CREATE DATABASE IF NOT EXISTS `excel` /*!40100 DEFAULT CHARACTER SET utf8 */;
USE `excel`;

-- Дамп структуры для таблица excel.excel_data
CREATE TABLE IF NOT EXISTS `excel_data` (
  `data` date NOT NULL,
  `name` varchar(50) NOT NULL,
  `age` int(11) NOT NULL,
  `login` varchar(50) NOT NULL,
  `balans` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Дамп данных таблицы excel.excel_data: ~3 rows (приблизительно)
/*!40000 ALTER TABLE `excel_data` DISABLE KEYS */;
INSERT INTO `excel_data` (`data`, `name`, `age`, `login`, `balans`) VALUES
	('2018-08-27', 'Alex', 1, '1', 1),
	('2018-08-27', 'Bro', 2, '2', 2),
	('2018-08-27', 'You', 3, '3', 4);
/*!40000 ALTER TABLE `excel_data` ENABLE KEYS */;

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IF(@OLD_FOREIGN_KEY_CHECKS IS NULL, 1, @OLD_FOREIGN_KEY_CHECKS) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
