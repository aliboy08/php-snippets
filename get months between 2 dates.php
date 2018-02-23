<?php
$start = new DateTime('2013-08-01');
$start->modify('first day of this month');
$end = new DateTime(date('Y-m-d')); // current date
$end->modify('first day of next month');
$interval = DateInterval::createFromDateString('1 month');
$period = new DatePeriod($start, $interval, $end);

foreach ($period as $dt) {
	echo $dt->format("F Y") . '<br/>';
}