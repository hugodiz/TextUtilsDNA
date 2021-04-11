# MaxHeapMechanics notes:
(a "Best K matches" Priority Queue implementation) 
by Hugo Diz, 2021

This module implements a data structure for holding the "best K results so far", in the context of, given an input text A(h), scanning a (possibly 2D) array B row by row, testing the quality of a match between A(h) and each such B(m,n), then storing the critical info of that B(m,n) in the BestK max heap (if indeed it is among the best K matches so far).

The foundation of this technique is the well-known method of using a binary tree with an imposed "heap property" (ie. an hierarchy between nodes which is enforced all the time by occasionaly required nodes to switch contents). In our case, we define and tweak the specifics of these relationships and the acts of "enqueueing and dequeueing", in order to fit our purpose of keeping a record of the "best K matches so far", as well as a way to ensure we always efficiently get a sorted list of of the heap, once LSDLOOKUP is done scanning the alookup_array.

The match value of any given B(m,n) (ie. its priority) is represented by a triplet of integers   
(Levenshtein Distance to A(h), m, n)   

That is, the row and column positions of B(m,n) in B (ie. [m,n]) are featured in the triplet and are part of the value.   
This is because we must not admit the possibility of a tie between two B(m,n)'s, seeing as we want LSDLOOKUP to be stable in the sense that it always resolves ties in the same predictable way.

Note that throughout this algorithm, "better" match == LOWER priority : the "worst match thus far" has the highest priority.

The spec of LSDLOOKUP says it returns the first K occurrences of the least Levenshtein distances found. We also specify that the B array is scanned row by row. So:
- If the score (Lev dist) of a match is lower, that match is better, and hence it has lower priority
- The score being equal, a match with lower m came first (upper row) (regardless of n), hence the lower m match is better
- All else being equal (same score, same row), a match with lower n is better, because rows are scanned left-to-right.

This suggests a natural way to compare triplets, implemented in the "Precedes" function: "precedes" means "it's a better match".

Under the hood, our heap is just an array with K+1 entries, BestK[x], where x = 0..K ; BestK[0] is never used, just ignore it
Let's say K = 6. Here's the heap structure
 
                                                  BestK[1]
                                                 /        \
                                         BestK[2]          BestK[3]
                                        /     \              /     
                                BestK[4]    BestK[5]    BestK[6]  

Our array becomes a binary tree by simply defining filial relationships between nodes on the basis of their index positions. Specifically:

- Every node can have zero, one or two children. The highest possible index is K = 6. The root node is BestK[1]
- A node's first (ie. left) or *only* child always has position Child1Pos = ParentPos * 2
    - Therefore, ParentPos * 2 > K is equivalent to "Parent not having children / node not being a parent"
        - In the example, this applies to nodes 4, 5 and 6
        - Also in the example, nodes 1, 2, 3 have children
- A node's second(right) child always has position = Child2Pos = ParentPos * 2 + 1
    - Therefore, ParentPos * 2 + 1 > K is equivalent to "Parent not having a second child"
        - In the example, this applies to node 3

- A corolary to this is that if ParentPos * 2 = K, this means simultaneously that:
    - node has children
    - node does not have a second child
        - Therefore, this parent has a single child
    - As stated already, in the example, this applies to node 3

The heap starts by being initialized with all nodes = (Threshold, -1, -1)
    Threshold = typo_tolerance + 1, where typo_tolerance is the user-specified maximum allowed Levenshtein Distance
    This means that any match at all which is not rejected a priori due to not respecting the threshold,
        will be "better" than (Threshold, -1, -1), because it's Lev dist will be < Threshold
    For this reason, we call (Threshold, -1, -1) entries "idles", because anything can overtake them

The heap works by 2 actions: try inserting (enqueueing) a node, and dequeuing the highest priority node

## ENQUEUE:
    The heap always maintains the following property, whenenver something is enqueued or dequeued:
        A parent is always a worse match (or equal, although there are no real ties here) than its children
    So when trying to enqueued a node, the node is tested against BestK[1] first:
        if the candidate node is not better than BestK[1], nothing is inserted
        if the candidate is better than BestK[1], throw away the contents of BestK[1] and replace with candidate node
    If the candidate node was inserted in BestK[1], now we need to re-establish the heap property:
        If the candidate (now in BestK[1]) is better than either child, then
            switch places with the worst child (or only child, or left child if all else is equal)
        If that candidate is in a position that has no children, or is not better than either child, do nothing
    The above check, based on the candidate initially being in BestK[1] will have either resulted in
        - candidate switching places with BestK[2]
        - candidate switching places with BestK[3]
        - candidate staying in BestK[1]
        If candidate moved (into either BestK[2] or BestK[3]), then repeat the check
            (this time based on the candidate initially being on BestK[2] or BestK[3] and potentially moving further)
        If candidate hasn't moved, then it's settled, the heap property is guaranteed to hold
    Note how the insertion of a node always takes a logarithmic(K) number of operations, in terms of time complexity

## DEQUEUE:
By the end of scanning B, whatever the heap holds are the best K matches found for our A(h). How to retrieve them?
    Since we want to extract each of them by order, so we can fill the Output(h, ..) rowand present by order,
        we dequeue each item sequentially, each time extracting the highest priority element (ie. the worst match)
            and then we keep filling the Output row "back to front"

The way to dequeue the the heap (essentially always forcing the root node out, since the root is always the worst)
    is to drop an "anchor"; an anchor is the entry (-1,-1,-1), which, when enqueued, always "sinks" to the bottom,
        because the anchor is technically a better match than any possible real match (it has Lev dist = -1)
    Hence, the anchor is guaranteed to overtake the root node.
    The important difference here, compared to the normal Enqueue action, is that we SAVE the contents of the
        root node, which get thrown away when the anchor takes its place. We give those contents to the Output array
    After the anchor taking over the root node, the heap reacts by enforcing the property:
        we know that reaction ends with the anchor sinking to a childless node;
        we know that, by the end of the reaction, the next-worst match will have necessarily made its way to the root
    As a result, we can keep dropping anchors sequentially, until we've dropped K in total:
        By then, the entire heap will be filled with anchors, and the real contents of the heap
            will have been dequeued by order of worst match to best
