using System.Text;

public class Trie
{
    private Node root = new Node();

    public void Insert(string key)
    {
        StringBuilder @string = new StringBuilder(key);
        Node current = root;
        while (@string.Length > 0)
        {
            if (current.children[@string[0] - 'a'] == null)
                current.children[@string[0] - 'a'] = new Node()
                {
                    WordEnd = @string.Length == 1
                };

            current = current.children[@string[0] - 'a'];
            @string.Remove(0, 1);
        }
    }

    private class Node
    {
        public Node[] children = new Node[26];
        public bool WordEnd = false;
    }
}
