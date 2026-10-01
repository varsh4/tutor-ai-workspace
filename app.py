
896
897
898
899
900
901
902
903
904
905
906
907
908
909
910
911
912
913
914
915
916
917
918
919
920
921
922
923
924
925
926
927
928
929
930
931
932
933
934
935
936
937
938
939
940
941
942
943
944
945
946
947
948
949
950

        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.session_state.messages:
        if st.button("🗑️ Clear conversation", use_container_width=True):
            st.session_state.messages = []
            st.rerun()

    if st.session_state.messages:
        st.markdown("### 🎨 Visual help")
        if st.button("Create visual for last student question", use_container_width=True):
            last_student = next(
                (
                    m["content"]
                    for m in reversed(st.session_state.messages)
                    if m["role"] == "student"
                ),
                "",
            )
            if last_student:
                with st.spinner("Building a colorful teaching visual…"):
                    image_bytes, error = generate_visual(
                        "Create the most useful educational visual for this student question. "
                        "If mathematical, prioritize an accurate graph/coordinate visual; "
                        "otherwise create the clearest colorful diagram or infographic.\n\n"
                        + last_student
                    )
                if image_bytes:
                    st.session_state.generated_visuals.append({
                        "prompt": last_student,
                        "bytes": image_bytes,
                    })
                    st.rerun()
                else:
                    st.error(error)

    st.markdown(
        """
        <div class="panel-card">
            <div class="panel-title">✨ Tutoring behavior</div>
            <div class="small-muted">
                • Independent verification<br>
                • Student-work assessment<br>
                • Step-by-step teaching<br>
                • Alternate explanations<br>
                • Calculations & data checks<br>
                • File-based assessment<br>
                • Research support
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
