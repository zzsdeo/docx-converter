package main

type Spec struct {
	system string
	rows   []Row
}

type Row struct {
	name       string
	partNumber string
	qty        float64
	measure    string
}

type RS Spec
