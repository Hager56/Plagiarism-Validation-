using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // For xlsx format
using NPOI.HSSF.UserModel;
using static Program;
using static NPOI.HSSF.Util.HSSFColor;
using System.Diagnostics;
class Program
{

    public class Data
    {
        public string? path;
        public string? my_cell;
        public int similarity;
        public int similarity1;
        public int similarity2;
        public int? lines_matched;
    }

    public static int ExtractSimilarityFromString(string? str)
    {
        Regex regex = new Regex(@"\d+");
        if (str == null)
        {
            throw new Exception("No integer found in the input string.");
        }
        Match match = regex.Match(str);

        if (match.Success)
        {
            //Console.WriteLine($"Extracted Similarity: {match.Value}"); // Print extracted similarity for debugging
            return int.Parse(match.Value);
        }
        throw new Exception("No integer found in the input string.");
    }
    public static List<Tuple<Data, Data>> ReadExcel(string filePath)
    {
        IWorkbook workbook;
        List<Tuple<Data, Data>> readableSheet;
        using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            if (filePath.EndsWith(".xlsx"))
                workbook = new XSSFWorkbook(file);
            else if (filePath.EndsWith(".xls"))
                workbook = new HSSFWorkbook(file);
            else
                throw new Exception("Invalid file extension");

            ISheet sheet = workbook.GetSheetAt(0); // Assuming you want the first sheet

            readableSheet = new List<Tuple<Data, Data>>();
            for (int rowIdx = 1; rowIdx <= sheet.LastRowNum; rowIdx++)
            {
                IRow row = sheet.GetRow(rowIdx);
                if (row != null)
                {
                    // fetch first path 
                    ICell stCell = row.GetCell(0);
                    string? stCellData = stCell.ToString();
                    string? stPath = stCellData?.Split('(')[0];
                    int stIndex = ExtractSimilarityFromString(stCellData?.Split('(')[0]);
                    string? stSimilarity = stCellData?.Split('(')[1];
                    int stSimValue1 = ExtractSimilarityFromString(stCellData?.Split('(')[1]); // Corrected line

                    int stSimValue = ExtractSimilarityFromString(stSimilarity);

                    string? stLinesMatched = row.GetCell(2)?.ToString();
                    Data st = new();
                    st.similarity = stSimValue;
                    st.path = stPath;
                    st.my_cell = stCellData;
                    st.similarity1 = stSimValue1; // Corrected line
                    st.lines_matched = int.Parse(stLinesMatched);

                    // fetch second path
                    ICell ndCell = row.GetCell(1);
                    string? ndCellData = ndCell.ToString();
                    string? ndPath = ndCellData?.Split('(')[0];
                    int ndIndex = ExtractSimilarityFromString(ndCellData?.Split('(')[0]);
                    int ndSimValue2 = ExtractSimilarityFromString(ndCellData?.Split('(')[1]);

                    string? ndSimilarity = ndCellData?.Split('(')[1];
                    int ndSimValue = ExtractSimilarityFromString(ndSimilarity);
                    string? ndLinesMatched = row.GetCell(2)?.ToString();

                    Data nd = new();
                    nd.similarity = ndSimValue;
                    nd.path = ndPath;
                    nd.similarity1 = stSimValue;
                    nd.similarity2 = ndSimValue;
                    int temp = st.similarity;
                    nd.my_cell = ndCellData;
                    nd.lines_matched = int.Parse(ndLinesMatched);
                    //st.similarity += nd.similarity;
                    //nd.similarity += temp;
                    st.similarity = Math.Max(st.similarity, nd.similarity);
                    nd.similarity = Math.Max(st.similarity, nd.similarity);
                    readableSheet.Add(Tuple.Create(st, nd));
                }
            }
            return readableSheet;
        }
    }

    public struct dataOfEdge
    {
        public string destination;
        public string source;
        public string my_cell1;
        public string my_cell2;
        public int similarity1;
        public int similarity2;
        public int simillarity;
        public int linesMatch;
    }


    public struct GraphData
    {
        public List<string> vertecis;
        public Dictionary<string, List<dataOfEdge>> adj_list;
    }

    // f1 -> { (f2,13) }

    //Making undirected adj_list of the given sheet
    static GraphData construct_graph(List<Tuple<Data, Data>> sheet)
    {
        Dictionary<string, List<dataOfEdge>> adj_list = new Dictionary<string, List<dataOfEdge>>();
        List<string> vertices = new List<string>();
        foreach (var row in sheet)
        {
            if (!adj_list.ContainsKey(row.Item1.path))
            {
                adj_list.Add(row.Item1.path, new List<dataOfEdge>());
                vertices.Add(row.Item1.path);
            }
            if (!adj_list.ContainsKey(row.Item2.path))
            {
                adj_list.Add(row.Item2.path, new List<dataOfEdge>());
                vertices.Add(row.Item2.path);
            }
            //storing source and destination and cells and similarity of each row in (edgeData struct)
            adj_list[row.Item1.path].Add(new dataOfEdge
            {

                destination = row.Item2.path,
                source = row.Item1.path,
                my_cell1 = row.Item1.my_cell,
                my_cell2 = row.Item2.my_cell,
                similarity1 = row.Item1.similarity1,
                similarity2 = row.Item2.similarity2,
                simillarity = row.Item1.similarity,
                linesMatch = (int)row.Item1.lines_matched

            });
            adj_list[row.Item2.path].Add(new dataOfEdge
            {

                destination = row.Item1.path,
                source = row.Item2.path,
                my_cell1 = row.Item1.my_cell,
                similarity1 = row.Item1.similarity1,
                similarity2 = row.Item2.similarity2,
                simillarity = row.Item1.similarity,
                my_cell2 = row.Item2.my_cell,
                linesMatch = (int)row.Item2.lines_matched

            });


        }
        GraphData data = new GraphData();
        data.adj_list = adj_list;
        data.vertecis = vertices;
        return data;


    }
    public class DisjointSet
    {
        private Dictionary<string, string> parent = new Dictionary<string, string>();

        public void MakeSet(string vertex)
        {
            parent[vertex] = vertex;
        }

        public string FindSet(string vertex)
        {
            if (parent[vertex] != vertex)
                parent[vertex] = FindSet(parent[vertex]);
            return parent[vertex];
        }

        public void Union(string vertex1, string vertex2)
        {
            parent[FindSet(vertex1)] = FindSet(vertex2);
        }
    }


    public static List<dataOfEdge> FindMaxSpanningTree(List<string> group, List<dataOfEdge> edgesOfEachGroup)
    {
        DisjointSet disjointSet = new DisjointSet();
        List<dataOfEdge> MSTList = new List<dataOfEdge>();

        foreach (var vertex in group)
        {
            disjointSet.MakeSet(vertex);
        }

        var sortedEdges = edgesOfEachGroup.OrderByDescending(e => e.simillarity).ThenByDescending(e=>e.linesMatch);

        dataOfEdge last_1_Edge = new dataOfEdge { simillarity = -1, linesMatch = 0 };
        string temp="";

        foreach (var edge in sortedEdges)
        {
            string u = edge.source;
            string v = edge.destination;

            if (disjointSet.FindSet(u) != disjointSet.FindSet(v))
            {

                MSTList.Add(edge);
                disjointSet.Union(u, v);

                last_1_Edge = edge;
                temp = v;
            }
            

        }

        var sortedByLineMatch = MSTList.OrderByDescending(e => e.linesMatch).ToList();

        return sortedByLineMatch;
    }

    public struct Connect
    {
        public List<string> verteces;
        public List<dataOfEdge> edgesOfEachGroup;

    }
    //making BFS on given root and return it's connected graph  {vertecies ,edge_list}
    public static Connect DetectConnectivity(Dictionary<string, List<dataOfEdge>> adj_list, string root, Dictionary<string, bool> visit, Dictionary<string, string> color)
    {

        //Dictionary<string, string> color = new Dictionary<string, string>();

        List<dataOfEdge> edgesOfEachGroup = new List<dataOfEdge>();

        Queue<string> Q = new Queue<string>();

        List<string> elements = new List<string>();

        color[root] = "grey";
        visit[root] = true;
        Q.Enqueue(root);

        while (Q.Count != 0)
        {
            string vertex = Q.Dequeue();

            List<dataOfEdge> neighbors = adj_list[vertex];

            foreach (var n in neighbors)
            {

                string neighborPath = n.destination;

                if (color[neighborPath] == "white")
                {
                    color[neighborPath] = "grey";

                    Q.Enqueue(neighborPath);
                }
                if ( color[neighborPath] == "grey")
                {
                    edgesOfEachGroup.Add(n);

                }
            }

            color[vertex] = "Black";

            elements.Add(vertex);
            visit[vertex] = true;
        }
        Connect c = new Connect();
        c.verteces = elements;
        c.edgesOfEachGroup = edgesOfEachGroup;
        return c;

    }
    // BFS all vertecies in the grapg to detect connectivity between
    public static void BFSAll(List<string> vertecis, Dictionary<string, List<dataOfEdge>> adj_list)
    {

        Dictionary<string, bool> visited = new Dictionary<string, bool>();
        Dictionary<string, string> color = new Dictionary<string, string>();
        //initiate 
        foreach (var ver in vertecis)
        {
            visited[ver] = false;
            color[ver] = "white";
        }

        List<dataOfEdge> newEdges = new List<dataOfEdge>();

        Connect n = new Connect();

        string filePath = @"C:\Users\Alrahma\Desktop\minee\export.csv";

        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
        using (StreamWriter writer = new StreamWriter(filePath))
        {
            writer.WriteLine("File 1,File 2,Line Matches");

            foreach (var v in vertecis)
            {
                if (visited[v] == false)
                {
                    //takes a vertex and make BFS on to return it's  graph { connected vertices , adj_list}
                    //mark all  connected vertcies of (n) as TRUE
                    n = DetectConnectivity(adj_list, v, visited, color); // ex : return (1,2,3) and edges between 

                    //making spaning tree on every connected group of vertcies and return edges after deleting the redundant edge
                    newEdges = FindMaxSpanningTree(n.verteces, n.edgesOfEachGroup);
                    foreach (var edge in newEdges)
                    {
                        //Console.WriteLine($"  {edge.my_cell1}       {edge.my_cell2}        {edge.linesMatch}");
                        string line = $"{edge.my_cell1},{edge.my_cell2},{edge.linesMatch}";
                        writer.WriteLine(line);
                    }
                    //Console.WriteLine();
                   // writer.WriteLine();
                }
            }
        }
    }

    //my code
    public static Connect DetectConnectivity2(Dictionary<string, List<dataOfEdge>> adj_list, string root)
    {
        var edgesOfEachGroup = new List<dataOfEdge>();
        var visited = new Dictionary<string, bool>();
        var color = new Dictionary<string, string>();
        var elements = new List<string>();

        foreach (var ver in adj_list.Keys)
        {
            visited[ver] = false;
            color[ver] = "white";
        }

        var Q = new Queue<string>();
        color[root] = "grey";
        visited[root] = true;
        Q.Enqueue(root);

        while (Q.Count != 0)
        {
            string vertex = Q.Dequeue();
            List<dataOfEdge> neighbors = adj_list[vertex];

            foreach (var n in neighbors)
            {
                string neighborPath = n.destination;
                if (color[neighborPath] == "white")
                {
                    color[neighborPath] = "grey";
                    Q.Enqueue(neighborPath);
                }
                if (color[neighborPath] == "white" || color[neighborPath] == "grey")
                {
                    edgesOfEachGroup.Add(n);
                }
            }
            color[vertex] = "Black";
            elements.Add(vertex);
            visited[vertex] = true;
        }

        Connect c = new Connect();
        c.verteces = elements;
        c.edgesOfEachGroup = edgesOfEachGroup;
        return c;
    }

    public static List<Tuple<List<string>, double, int>> FindComponentsAndAvgSimilarity(GraphData graphData)
    {
        Dictionary<string, bool> visited = new Dictionary<string, bool>();
        Dictionary<string, string> color = new Dictionary<string, string>();
        var components = new List<Connect>();
        //var visited = new HashSet<string>();
        foreach (var ver in graphData.vertecis)
        {
            visited[ver] = false;
            color[ver] = "white";
        }
        foreach (var chk in graphData.adj_list)
        {


            var root = chk.Key;
            if (visited[root] == false)
            {
                //var component = DetectConnectivity2(graphData.adj_list, root);
                var component = DetectConnectivity(graphData.adj_list, root, visited, color);
                components.Add(component);
                foreach (var vertex in component.verteces)
                {
                    visited[vertex] = true;
                }
            }
        }

        var result = new List<Tuple<List<string>, double, int>>();

        foreach (var component in components)
        {
            var edgesOfEachGroup = component.edgesOfEachGroup;

            var uniqueVertices = new HashSet<string>();
            foreach (var edge in edgesOfEachGroup)
            {
                uniqueVertices.Add(edge.source);
                uniqueVertices.Add(edge.destination);
            }
            int verticesCount = (edgesOfEachGroup.Count) * 2; //number of vertices(kol edge*2)

            double totalSimilarity = edgesOfEachGroup.Sum(edge => edge.similarity1 + edge.similarity2);
            double avgSimilarity = (totalSimilarity / verticesCount);
            avgSimilarity = Math.Round(avgSimilarity, 1);//round

            edgesOfEachGroup.Sort((x, y) => Math.Max(y.similarity1, y.similarity2).CompareTo(Math.Max(x.similarity1, x.similarity2)));


            var numericVertices = component.verteces.Select(vertex => Regex.Match(vertex, @"\d+").Value).Distinct().ToList();
            result.Add(new Tuple<List<string>, double, int>(numericVertices, avgSimilarity, verticesCount));
            result.Sort((x, y) => y.Item2.CompareTo(x.Item2));

        }

        return result;
    }



    static void Main(String[] args)
    {

        List<Tuple<Data, Data>> sheet;
        sheet = ReadExcel("C:\\Users\\Alrahma\\Desktop\\2-Input.xlsx");
        GraphData graphData = construct_graph(sheet);
        List<string> vertecis = graphData.vertecis;
        //start
        Stopwatch stopwatch1 = new Stopwatch();
        Stopwatch stopwatch2 = new Stopwatch();

        Stopwatch stopwatch3 = new Stopwatch();

        stopwatch3.Start();

        stopwatch1.Start();

        var components = FindComponentsAndAvgSimilarity(graphData);
        Console.WriteLine("FIRST: Groups Stats\r\n\n");
        string filePath = @"C:\Users\Alrahma\Desktop\minee\export-stat.csv";

        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
        using (StreamWriter writer = new StreamWriter(filePath))
        {
            writer.WriteLine("File 1,File 2,Line Matches");
            //Console.WriteLine("Component Index\t- Vertices\t\t\t - Average Similarity\t\t - Component Count");
            writer.WriteLine("Component Index,Vertices,Average Similarity,Component Count");
            int componentIndex = 1;
            foreach (var component in components)
            {
                var sortedVertices = component.Item1.OrderBy(vertex => int.Parse(Regex.Match(vertex, @"\d+").Value)).ToList();
                double avgSimilarity = component.Item2;
                int componentCount = component.Item1.Count;
               // Console.WriteLine($"{componentIndex} \t\t {string.Join(", ", sortedVertices)} \t\t\t {avgSimilarity} \t\t\t {componentCount}");
                writer.WriteLine($"{componentIndex},{string.Join(" - ", sortedVertices)},{avgSimilarity},{componentCount}");
                componentIndex++;
            }
        }
        stopwatch1.Stop();
        long elapsedMilliseconds3 = stopwatch1.ElapsedMilliseconds;

        Console.WriteLine("Elapsed time stat : " + elapsedMilliseconds3 + " milliseconds");
        //stopwatch.Stop();
        //long elapsedMilliseconds = stopwatch.ElapsedMilliseconds;

        //// Print the elapsed time
        //Console.WriteLine("Elapsed time: " + elapsedMilliseconds + " milliseconds");
        //WriteStatisticsToFile(components);
        //Print the vertices
        //Console.WriteLine("\nVertices:");
        //foreach (var vertex in graphData.vertecis)
        //{
        //    Console.WriteLine(vertex);
        //}
        stopwatch2.Start();
        Dictionary<string, List<dataOfEdge>> adj_list = graphData.adj_list;

        //List<string> vertecis = graphData.vertecis;

        string root = vertecis[0];


        Console.WriteLine("\nSECOND: Refined Pairs(MST)\r\n\n");

        BFSAll(vertecis, adj_list);

        stopwatch2.Stop();
        long elapsedMilliseconds2 = stopwatch2.ElapsedMilliseconds;

        Console.WriteLine("Elapsed time MST : " + elapsedMilliseconds2 + " milliseconds");

        long elapsedMilliseconds = stopwatch3.ElapsedMilliseconds;

        // Print the elapsed time
        Console.WriteLine("\n total Elapsed time: " + elapsedMilliseconds + " milliseconds");
    }

}
