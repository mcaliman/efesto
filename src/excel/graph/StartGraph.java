/*
 * Efesto - Excel Formula Extractor System and Topological Ordering algorithm.
 * Copyright (C) 2017 Massimo Caliman mcaliman@caliman.biz
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as published
 * by the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
 * If AGPL Version 3.0 terms are incompatible with your use of
 * Efesto, alternative license terms are available from Massimo Caliman
 * please direct inquiries about Efesto licensing to mcaliman@caliman.biz
 */

package excel.graph;

import excel.grammar.Formula;
import excel.grammar.Start;
import excel.grammar.formula.functioncall.EXCEL_FUNCTION;
import excel.grammar.formula.functioncall.binary.Binary;
import excel.parser.StartList;

import java.util.*;

public class StartGraph {

    private final HashMap<Start, Node> graph;

    public StartGraph() {
        this.graph = new HashMap<>();
    }

    public void addNode(Start start) {
        if (start.isTerminal()) return;
        Node u = this.graph.get(start);
        if (u == null) {
            u = new Node(start);
            this.graph.put(start, u);
        } else if (notEquals(u, start)) {
            u.setValue(start);
            this.graph.put(start, u);
        }
    }

    public void addEdge(Start x, Start y) {
        if (x.isTerminal() || y.isTerminal()) return;
        if (x.getAddr().equalsIgnoreCase(y.getAddr())) return;
        Node u = graph.get(x);
        Node v = graph.get(y);
        Edge edge = new Edge(u, v);
        u.addEdge(edge);
    }

    public void add(Binary operation) {
        Start left = operation.getlFormula();
        Start right = operation.getrFormula();
        addNode(right);
        addNode(left);
        addNode(operation);
        addEdge(right, operation);
        addEdge(left, operation);
    }

    public void add(EXCEL_FUNCTION function) {
        Formula[] args = function.getArgs();
        for (Formula arg : args)
            addNode(arg);
        addNode(function);
        for (Formula arg : args)
            addEdge(arg, function);
    }

    /**
     * Use kahn Top Sort
     *
     * @return
     */
    public StartList topologicalSort() {
        StartList result = new StartList();
        Queue<Node> queue = new ArrayDeque<>();
        Collection<Node> nodes = this.graph.values();
        List<Edge> edges = this.edges();
        for (Node v : nodes)
            if (!hasIncomingEdges(v, edges))
                queue.add(v);
        while (!queue.isEmpty()) {
            Node v = queue.poll();
            result.add(v.value());
            List<Edge> outgoingEdges = outgoingEdges(v);
            for (Edge e : outgoingEdges) {
                Node s = e.src();
                Node t = e.dest();
                this.removeEdge(s.value(), t.value());
                Node end = e.dest();
                List<Edge> edges1 = this.edges();
                if (!hasIncomingEdges(end, edges1))
                    queue.add(end);
            }
        }
        if (!this.edges().isEmpty()) {
            System.err.println("error when sort!. this.edges().size()=" + this.edges().size());
            return result;
        }
        return result;
    }

    private void removeEdge(Start x, Start y) {
        Node u = graph.get(x);
        Node v = graph.get(y);
        u.removeEdgeTo(v);
    }

    private List<Edge> edges() {
        List<Edge> results = new ArrayList<>();
        Collection<Node> nodes = this.graph.values();
        for (Node node : nodes)
            results.addAll(node.edges());
        return results;
    }

    private boolean hasIncomingEdges(Node v, List<Edge> allEdges) {
        for (Edge edge : allEdges)
            if (edge.dest().equals(v)) return true;
        return false;
    }

    private List<Edge> outgoingEdges(Node v) {
        List<Edge> outgoingEdges = new ArrayList<>();
        List<Edge> edges = edges();
        edges.stream().filter((edge) -> (edge.src().equals(v))).forEachOrdered(outgoingEdges::add);
        return outgoingEdges;
    }

    private boolean notEquals(Node u, Start start) {
        return u != null && !u.value().toString().equals(start.toString());
    }

}
