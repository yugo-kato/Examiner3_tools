use std::fs::File;
use std::io::{BufRead, BufReader, Write};
use regex::Regex;
use calamine::{DataType, Reader, open_workbook, Xlsx};
use std::io;
use std::path::PathBuf;
use rfd::{FileDialog};

fn txt_to_latloncsv(openfilename: &str, outputfilename: &str) {
    let f = File::open(openfilename).expect("Failed to open input file");
    let mut out = File::create(outputfilename).expect("Failed to create output file");

    writeln!(out, "X(Scan ID.),Y(Channel No.),Latitude,Longitude").expect("Failed to write to output file");

    let reader = BufReader::new(f);
    let mut lines = reader.lines();

    for _ in 0..2 {
        lines.next();
    }

    let re = Regex::new(r"\d+").unwrap();
    if let Some(Ok(line)) = lines.next() {
        let result: Vec<&str> = re.find_iter(&line).map(|mat| mat.as_str()).collect();
        if result.len() == 3 {
            let nx = result[0].parse::<i32>().unwrap();
            let ny = result[1].parse::<i32>().unwrap();

            println!("{}, {}", nx, ny);

            for _ in 0..3 {
                lines.next();
            }

            for i in 0..nx {
                for j in 0..ny {
                    if let Some(Ok(line)) = lines.next() {
                        let parts: Vec<&str> = line.split('\t').collect();
                        writeln!(out, "{},{},{},{}", i, j, parts[0], parts[1]).expect("Failed to write to output file");
                    }
                }
            }
        }
    }
}

fn main() -> io::Result<()> {

    let input_file_path = match FileDialog::new().add_filter("Excel file",&["xlsx"]).pick_file() {
        Some(path) => path,
        None => return Err(io::Error::new(io::ErrorKind::Other, "No file selected.")),
    };

    let mut df: Vec<(String, i32, i32)> = vec![];

    // Load the Excel file
    let mut workbook: Xlsx<_> = open_workbook(input_file_path.clone()).expect("Failed to open Excel file");

    // Select the first sheet and read the data
    if let Some(Ok(r)) = workbook.worksheet_range_at(0) {
        for row in r.rows().skip(1) {
            let name = row[0].get_string().unwrap().to_string();
            let start_x = row[2].get_float().unwrap() as i32;
            let end_x = row[3].get_float().unwrap() as i32;
            df.push((name, start_x, end_x));
        }
    }

    for (i, (name, _, _)) in df.iter().enumerate() {
        println!("{}", i);

        let mut str_path = PathBuf::from(input_file_path.parent().unwrap());
        str_path.push(format!(r"{} - Region1.txt", name));

        //let openfilename = format!(r"\\ims\J\20240621とりまとめ\ASCII\{} - Region1.txt", name);
        let outputfilename = format!("{}_lat_lon.csv", &str_path.to_str().unwrap()[..&str_path.to_str().unwrap().len()-4]);
        txt_to_latloncsv(&str_path.to_str().unwrap(), &outputfilename);
    }

    Ok(())

}
