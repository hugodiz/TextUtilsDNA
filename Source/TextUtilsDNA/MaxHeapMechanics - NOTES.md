# MaxHeapMechanics notes:
(a "Best K matches" Priority Queue implementation)  
by Hugo Diz, 2021

This module implements a data structure for holding the "best K results so far", in the context of scanning a (possibly 2D) array B row by row (given an input text A(h)), testing the quality of the match between A(h) and each such B(m,n), then storing the critical info of B(m,n) in the BestK max heap (if indeed it sits among the best K matches so far).

The foundation of this technique is the well-known method of using a binary tree with an imposed "heap property" (ie. an hierarchy between nodes which is enforced all the time by requiring nodes to switch contents where needed). In our case, we define and tweak the specifics of these relationships and the acts of "enqueueing" and "dequeueing", in order to fit our purpose of keeping a record of the "best K matches so far", as well as a way to ensure we always efficiently get a sorted list out of the heap, after LSDLOOKUP is done scanning the lookup_array B.

The match value of any given B(m,n) (ie. its priority) is represented by a triplet of integers   
**(** Levenshtein Distance to A(h) **,** m **,** n **)**   

That is, the row and column positions of B(m,n) in B (ie. [m,n]) are featured in the triplet and are part of the value.   
This is because we must not admit the possibility of a tie between two B(m,n)'s, seeing as we want LSDLOOKUP to be stable in the sense that it always resolves ties in the same predictable way.

Note that throughout this algorithm, "better" match == LOWER priority : the "worst match thus far" has the highest priority.

The spec of LSDLOOKUP says it returns the first K occurrences of the least Levenshtein distances found. We also specify that the B array is scanned row by row. So:
- If the score (Lev dist) of a match is lower, that match is better, and hence it has lower priority
- The score being equal, a match with lower m came first (upper row) (regardless of n), hence the lower m match is better
- All else being equal (same score, same row), a match with lower n is better, because rows are scanned left-to-right.

This suggests a natural way to compare triplets, implemented in the "Precedes" function: "precedes" means "it's a better match".

Under the hood, our heap is just an array with K+1 entries, BestK[x], where x = 0..K ; BestK[0] is never used, we can just ignore it.

Let's say K = 6. Here's the heap structure
 
                                                  BestK[1]
                                                 /        \
                                         BestK[2]          BestK[3]
                                        /     \              /     
                                BestK[4]    BestK[5]    BestK[6]  

Our array becomes a binary tree by simply defining filial relationships between nodes on the basis of their index positions. Specifically:

- Every node can have zero, one or two children. The highest possible index in the example is K = 6. The root node is BestK[1].  
- A node's first (ie. left) or *only* child always has position Child1Pos = ParentPos * 2
    - Therefore, ParentPos * 2 > K is equivalent to "Parent not having children" ie. "node not being a parent"
        - In the example, this applies to nodes 4, 5 and 6
        - Also in the example, nodes 1, 2, 3 all have at least one child, ie. they all have children
- A node's second (ie. right) child always has position = Child2Pos = ParentPos * 2 + 1
    - Therefore, ParentPos * 2 + 1 > K is equivalent to "Parent not having a second child" ie. "Parent having at most one child"
        - In the example, this applies to node 3

- A corolary to this is that if ParentPos * 2 = K, this means simultaneously that:
    - node has children
    - node does not have a second child
        - Therefore, this parent has a single child
    - As stated already, in the example, this applies to node 3

The heap starts by being initialized with all nodes = (Threshold, -1, -1), where Threshold = *typo_tolerance* + 1 and where *typo_tolerance* is the user-specified maximum allowed Levenshtein Distance (where technically *typo_tolerance* can't be set to higher than 2,147,483,646 = MaximumInteger - 1). 

Since the size of a string held in an Excel cell is capped at a value orders of magnitude below this threshold, and the Levenshtein Distance between two strings cannot ever be greater than the size of the longest string, than this means we can be sure that any match *at all* (which is not rejected a priori due to not respecting the threshold) will always be trivially "better" than (Threshold, -1, -1), no matter what. Any match which might ever attempt to enter the best K (ie. not rejected a prior), will necessarily have a Levenshtein distance < Threshold by definition.

For this reason, we call (Threshold, -1, -1) entries "balloons", because any valid match at all will overtake them, ie. will be "better than nothing". The "balloons" use "-1" for the positions in order for us to easily identify them as "actually non-matches". This is safe because although a row postiion of -1 would technically be better than the position of any real match, a tie-break which would make use of that fact can't actually ever come to pass, because the Levenshtein Distance score of a "ballon" will always be strictly higher than the one of any real match against which it may come to be compared. In other words, a ballon will always float upwards in the tree when compared against *anything real* (to see exactly how, see the next section, **Enqueue**).

The heap works by 2 actions: try inserting (enqueueing) a node, and extracting (dequeuing) the highest priority node. Interestingly, the act of extraction works by inserting a dummy value (an "anchor") which is the functional antithesis of the "ballon".

## Enqueue:
The heap always maintains the following property, whenever something is enqueued or dequeued:
- A parent is always a worse match (or equal, although there are no real ties here) than its children
- So, when trying to enqueue a node, that "candidate" node is tested against BestK[1] first:
    - if the candidate node is not better than BestK[1], nothing is inserted, candidate discarded
    - if the candidate is better than BestK[1], throw away the contents of BestK[1] and replace with candidate node
- If the candidate node was successfully inserted in BestK[1], now we need to re-establish the heap property:
    - If the candidate (now in BestK[1]) is better than either child, then switch places with the worst child (or only child, or left child if all else is equal)
    - If that candidate is in a position that has no children, or which is not better than either child, do nothing, candidate has settled
- The above check, based on the candidate initially being in BestK[1] will have either resulted in
    - candidate switching places with BestK[2]
    - candidate switching places with BestK[3]
    - candidate staying in BestK[1]
- If candidate moved (into either BestK[2] or BestK[3]), then repeat the check (this time based on the candidate initially being on BestK[2] or BestK[3] and potentially moving further)
- If candidate hasn't moved (which must happen eventually), then it's settled at that point and the heap property is guaranteed to hold
- Note how the insertion of a node always takes a logarithmic(K) number of operations, in terms of time complexity, since the worst-case number of operations required to settle a node is proportional to number of *generations / depth* of the heap, and not the number of nodes in the heap.

## Dequeue:
By the end of scanning B, whatever the heap holds are the best K matches found for our A(h). How to retrieve them in order?

Since we want to extract each of them by order, so we can fill the Output(h, ..) row and present results by best to worst and favouring "first found" entries when the Lev score is the same, we really just dequeue each item sequentially, making sure that each time we do it, we're extracting the highest priority element (ie. the worst match in the best K), and then we keep filling the Output row "back to front".

Our way to dequeue the heap (essentially always forcing the root node out, since the root is always the worst) is only adequate because:
- We will only dequeue everything at the end of the process
- We don't care about the heap state after we retrieve all best K by order

The gimmick here is dropping an "anchor"; an anchor is the entry (-1,-1,-1), which, when enqueued, always "sinks" to the bottom, because the anchor is technically a better match than any possible real match (because it has Lev dist = -1). Hence, the anchor is guaranteed to overtake the root node, forcing it out (ie. dequeueing it). There are 2 critical things this method seamlessly achieves (at the cost of shrinking the usable size of the heap with each anchor, of course):

- Since the root node is guaranteed to be forced out when an anchor is dropped, we can simply grab it as it comes out and hence we've retrieved a match (or possibly a ballon) from the heap
- Because we've really just enqueued a node, we can be sure that the same exact reaction mechanics as before will guarantee that the heap property holds, which means we can be sure that after the anchor has settled somewhere in the bottom, whatever sits at the root node is the next worst match.

The corollary to this is that by sequentially dropping K anchors, we are guaranteed a stream of extracted nodes which starts from the worst match (highest priority), ending in the best (lowest priority). These extracted nodes will comprise a row of Outputs for each lookup_value in LSDLOOKUP, shown best-to-worst left-to-right. If some extracted nodes are ballons, these will naturally come out before any real matches (because baloons float), and be converted to #N/A which will naturally form a padding on the right side of each affected Output row (which always has K entries, possibly with some or all #N/As). Of course, this also means the heap is now full of anchors and has become "unusable", because no real match would possibly have a place there. No problem, because we're done for this lookup_value and these anchors will be reset to ballons, come the next lookup_value.

Note how dropping K anchors takes K * log(k) time, because you insert a node (which is log(K)) K times. This time complexity is consistent with that of an efficient sorting algorithm (heap sort). Although the heap sort is technically not stable (ie. it does not always necessarily preserve the original order of tied values), we've worked around that problem by defining a triplet sorting rule which is guaranteed to never hold any actual ties. Or equivalently, we might say we introduced information for "explicitly remembering" the original order in the form of the 2nd and 3rd triplet elements. The reason why this is not wasteful is because we were always going to need to keep that "coordinate" information handy anyway, in order to actually lookup and output the final matched values, by the end of the function.

The beauty of this arrangement is that the iterative scanning of B (say it has M elements) costs M * log(K) in total. Then the sorting of the best K takes K * log(K). Because these are separate steps and you only sort once at the end, the total is M * log(K) + K * log(K) = (M + K) * log(K) -> M * log(K), considering that potentially M >> K, given that M can correspond to a very large dataset, and K can be no higher than 1024, according to the spec. 

A more naive approach might use a linear K-array (say best-to-worst == left-to-right) and just insert each new match in each proper place by (worst-case) scanning the entire K-array each time a new top-K-worthy match were found, and discarding the right-most element when needed. This would cost M * K time, which though not a tragedy, I think might be noticeably worse, except for quite small K's. 

So the difference between the naive M * K and our (M + K) * log(K) is due to the fact that we're not paying the cost of keeping our BestK holder sorted all the time - instead, we only bother to properly sort it once, at the end, which is all that's needed.

Of course, since the advantage of the heap grows with K, we could maybe argue that what's really "naive" here is not recognizing that our K won't tipically be very large, as it represents the number of output columns in Excel, so users are not likely to go crazy with it anyway. But they could, and the spec does allow for up to K = 1024. Plus, it was more fun doing it with the heap.
