class Delay {
	short(time := 0) {
		Sleep 200 + time
	}
	mid(time := 0) {
		Sleep 500 + time
	}
	long(time := 0) {
		Sleep 1000 + time
	}
}