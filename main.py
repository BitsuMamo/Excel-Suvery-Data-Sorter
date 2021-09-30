from SortData import SortData

if __name__ == "__main__":
    sort = SortData(input("Enter Path: "), input("Column of Num: "), input("Column X: "),input("Column Y: "),input("Column Elevation: "), input("Column Type: "), input("Column For X, Y, ELE Concat: "), input("Column For All Concat: "), float(input("X Bound: ")), float(input("Y Bound: ")))
    #sort = SortData("example.xlsx", "A", "B",300000,"C",700000,"D","E","F","G")
    sort.run()
